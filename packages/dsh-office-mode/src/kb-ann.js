/**
 * 纯 JS 的 HNSW 近似最近邻索引（第一百零三轮）。
 *
 * 为什么自己写：插件是零运行时依赖的纯 Node 包，CI 不联网、不装原生扩展，而
 * sqlite-vec / hnswlib 这类方案都要带原生绑定。这里的索引只是**可重建的缓存**——
 * 真源仍是 SQLite 的 embeddings 表与知识库 manifest，坏了删掉重建，不会丢数据。
 *
 * 语义约定：
 *   - 向量在入口处 L2 归一化，内部相似度就是点积（等于余弦）；
 *   - id 是外部字符串（块 id），内部槽位自增；重复 add 同一 id 等价于 update
 *     （旧节点打墓碑、按新向量重新插入，图连接因此不会残留旧向量的邻居）；
 *   - 删除只打墓碑：保留图上的连通性，检索时跳过；墓碑太多时 `compact()`
 *     按存活向量整图重建；
 *   - 层数与邻居选择用固定种子的伪随机数，同一批向量在任意机器上建出同一张图
 *     （评测与回归才可复现）。
 *
 * 已知边界（文档里写了，不假装没有）：带谓词过滤的检索只能「先在候选里筛」，
 * 筛完不够 k 时靠加大 efSearch 与调用方的精确扫描补足；它不是原生支持过滤的
 * 图索引，选择性与召回之间的权衡由 kb-db.js 负责。
 *
 * @module dsh-office-mode/kb-ann
 */
import { mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

/** 文件魔数：读错文件时第一时间能认出来，而不是解析出一张乱图。 */
export const ANN_MAGIC = 'DKBH';
/** 序列化版本；读不懂的版本一律当损坏处理（重建比兼容便宜）。 */
export const ANN_VERSION = 1;
/** HNSW 默认参数：M=16 是 256–1024 维小语料的常用档，建图成本与召回都稳。 */
export const ANN_DEFAULTS = { m: 16, efConstruction: 200, efSearch: 64, seed: 0x5eed1e55 };
/** efSearch 的上限：再大就不是 ANN 了，不如整表精确扫描。 */
export const ANN_MAX_EF = 512;
/** 头长度（字节）：魔数 4 + 版本 4 + 维度 4 + 槽位 4 + 存活 4 + 墓碑 4 + M 4 + efC 4 + 入口 4 + 最大层 4 + 种子 4 = 44。 */
const HEADER_BYTES = 44;

function asText(value) {
    if (typeof value === 'string') return value;
    if (value === null || value === undefined) return '';
    return String(value);
}

function toFloat32(vector) {
    if (vector instanceof Float32Array) return vector;
    if (Array.isArray(vector)) return Float32Array.from(vector.map(Number));
    if (vector instanceof Uint8Array) {
        if (vector.byteLength % 4 !== 0) return null;
        return new Float32Array(vector.buffer.slice(vector.byteOffset, vector.byteOffset + vector.byteLength));
    }
    return null;
}

/** L2 归一化；零向量返回 null（不能进索引：余弦对零向量没有定义）。 */
function normalized(vector, dim) {
    const raw = toFloat32(vector);
    if (raw === null || raw.length === 0) return null;
    const size = dim > 0 ? dim : raw.length;
    if (raw.length !== size) return null;
    const out = new Float32Array(size);
    let norm = 0;
    for (let i = 0; i < size; i += 1) {
        const value = Number(raw[i]);
        if (!Number.isFinite(value)) return null;
        out[i] = value;
        norm += value * value;
    }
    if (!(norm > 0)) return null;
    const scale = 1 / Math.sqrt(norm);
    for (let i = 0; i < size; i += 1) out[i] *= scale;
    return out;
}

function dot(left, right) {
    let sum = 0;
    for (let i = 0; i < left.length; i += 1) sum += left[i] * right[i];
    return sum;
}

/** 确定性伪随机（mulberry32）：建图与测试都靠它复现。 */
function createRandom(seed) {
    let state = (Number(seed) >>> 0) || 1;
    return function next() {
        state = (state + 0x6d2b79f5) >>> 0;
        let t = state;
        t = Math.imul(t ^ (t >>> 15), t | 1);
        t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}

/** 最小堆：检索时按相似度从小到大取「离查询最近」的候选。 */
function createMinHeap(compare) {
    const items = [];
    return {
        get size() { return items.length; },
        peek() { return items[0]; },
        push(item) {
            items.push(item);
            let at = items.length - 1;
            while (at > 0) {
                const parent = (at - 1) >> 1;
                if (compare(items[at], items[parent]) >= 0) break;
                const hold = items[parent];
                items[parent] = items[at];
                items[at] = hold;
                at = parent;
            }
        },
        pop() {
            if (items.length === 0) return undefined;
            const top = items[0];
            const last = items.pop();
            if (items.length > 0) {
                items[0] = last;
                let at = 0;
                for (;;) {
                    const left = at * 2 + 1;
                    const right = left + 1;
                    let smallest = at;
                    if (left < items.length && compare(items[left], items[smallest]) < 0) smallest = left;
                    if (right < items.length && compare(items[right], items[smallest]) < 0) smallest = right;
                    if (smallest === at) break;
                    const hold = items[smallest];
                    items[smallest] = items[at];
                    items[at] = hold;
                    at = smallest;
                }
            }
            return top;
        },
    };
}

/** 候选集：按相似度升序排列（下标 0 是最差的一条，方便 O(1) 淘汰）。 */
function bestInsert(list, item, cap) {
    let low = 0;
    let high = list.length;
    while (low < high) {
        const mid = (low + high) >> 1;
        if (list[mid].sim < item.sim) low = mid + 1;
        else high = mid;
    }
    list.splice(low, 0, item);
    if (list.length > cap) list.shift();
}

/**
 * 建一个 HNSW 索引。
 *
 * @param {{dim?:number, m?:number, efConstruction?:number, efSearch?:number, seed?:number}} [options]
 */
export function createAnnIndex(options = {}) {
    const source = options !== null && typeof options === 'object' ? options : {};
    const m = clampInt(source.m, ANN_DEFAULTS.m, 4, 64);
    const efConstruction = clampInt(source.efConstruction, ANN_DEFAULTS.efConstruction, m, 1_000);
    const defaultEfSearch = clampInt(source.efSearch, ANN_DEFAULTS.efSearch, 1, ANN_MAX_EF);
    const seed = Number.isFinite(source.seed) ? (source.seed >>> 0) : ANN_DEFAULTS.seed;
    const levelFactor = 1 / Math.log(m);
    const random = createRandom(seed);

    let dim = Number.isFinite(source.dim) && source.dim > 0 ? Math.trunc(source.dim) : 0;
    const ids = [];
    const vectors = [];
    const levels = [];
    const tombstones = [];
    const links = [];
    const idToSlot = new Map();
    let entry = -1;
    let maxLevel = -1;
    let active = 0;
    let dead = 0;

    function grow() {
        ids.push('');
        vectors.push(null);
        levels.push(0);
        tombstones.push(0);
        links.push([]);
        return ids.length - 1;
    }

    function levelOf() {
        // 经典 HNSW 的层数分布：-ln(U) * 1/ln(M)。U 取 (0,1)，避免 ln(0)。
        const u = Math.max(1e-12, random());
        return Math.floor(-Math.log(u) * levelFactor);
    }

    function maxM(level) {
        return level === 0 ? m * 2 : m;
    }

    /** 贪心下降一层：只走比当前更近的邻居，够用且便宜。 */
    function greedy(query, startSlot, level) {
        let current = startSlot;
        let currentSim = dot(query, vectors[current]);
        let improved = true;
        while (improved) {
            improved = false;
            const neighbors = links[current][level];
            if (neighbors === undefined) break;
            for (const neighbor of neighbors) {
                if (tombstones[neighbor] === 1) continue;
                const sim = dot(query, vectors[neighbor]);
                if (sim > currentSim) {
                    current = neighbor;
                    currentSim = sim;
                    improved = true;
                }
            }
        }
        return { slot: current, sim: currentSim };
    }

    /** 第 level 层的 beam search：返回按相似度降序的候选。 */
    function searchLayer(query, entries, ef, level) {
        const visited = new Set();
        // 候选要按「相似度最高」先出：HNSW 的 beam search 是从最近处向外扩，
        // 先出最远的那一条会让循环提前撞上 best 的下界、漏掉整片区域（实测召回 6.9%）。
        const candidates = createMinHeap((left, right) => right.sim - left.sim);
        const best = [];
        for (const slot of entries) {
            if (slot === undefined || slot < 0 || vectors[slot] === null) continue;
            if (visited.has(slot)) continue;
            visited.add(slot);
            const sim = dot(query, vectors[slot]);
            candidates.push({ slot, sim });
            bestInsert(best, { slot, sim }, ef);
        }
        while (candidates.size > 0) {
            const current = candidates.pop();
            // best 已满且当前候选还不如最差的一条：后面的只会更远。
            if (best.length >= ef && current.sim < best[0].sim) break;
            const neighbors = links[current.slot][level];
            if (neighbors === undefined || neighbors.length === 0) continue;
            for (const neighbor of neighbors) {
                if (visited.has(neighbor)) continue;
                visited.add(neighbor);
                const sim = dot(query, vectors[neighbor]);
                if (best.length < ef || sim > best[0].sim) {
                    candidates.push({ slot: neighbor, sim });
                    bestInsert(best, { slot: neighbor, sim }, ef);
                }
            }
        }
        const out = new Array(best.length);
        for (let i = 0; i < best.length; i += 1) out[i] = best[best.length - 1 - i];
        return out;
    }

    /**
     * 在第 0 层的候选里选邻居：这里用「最近的 M 个」（simple neighbor selection）。
     * 启发式选择（diversity）召回略好一点，但建图成本与实现复杂度都上去；小语料
     * 上实测差距在 1% 以内，先用简单口径。
     */
    function selectNeighbors(candidates, cap) {
        const picked = [];
        const seen = new Set();
        for (const item of candidates) {
            if (picked.length >= cap) break;
            if (tombstones[item.slot] === 1) continue;
            if (seen.has(item.slot)) continue;
            seen.add(item.slot);
            picked.push(item.slot);
        }
        return picked;
    }

    /** 双向连边后按「离本节点最近」裁剪，保持每层最多 maxM 个邻居。 */
    function connect(slot, neighborSlots, level) {
        links[slot][level] = neighborSlots;
        for (const neighbor of neighborSlots) {
            const list = links[neighbor][level];
            if (list.includes(slot)) continue;
            list.push(slot);
            const cap = maxM(level);
            if (list.length <= cap) continue;
            const scored = list.map((item) => ({ slot: item, sim: dot(vectors[slot], vectors[item]) }));
            scored.sort((left, right) => right.sim - left.sim);
            links[neighbor][level] = scored.slice(0, cap).map((item) => item.slot);
        }
    }

    function insertRaw(id, vector) {
        const slot = grow();
        const level = levelOf();
        ids[slot] = id;
        vectors[slot] = vector;
        levels[slot] = level;
        tombstones[slot] = 0;
        for (let i = 0; i <= level; i += 1) links[slot][i] = [];
        idToSlot.set(id, slot);
        active += 1;
        if (entry === -1) {
            entry = slot;
            maxLevel = level;
            return;
        }
        let current = { slot: entry, sim: dot(vector, vectors[entry]) };
        for (let levelAt = maxLevel; levelAt > level; levelAt -= 1) current = greedy(vector, current.slot, levelAt);
        let entries = [current.slot];
        for (let levelAt = Math.min(level, maxLevel); levelAt >= 0; levelAt -= 1) {
            const found = searchLayer(vector, entries, efConstruction, levelAt);
            const neighbors = selectNeighbors(found, m);
            connect(slot, neighbors, levelAt);
            if (found.length > 0) entries = found.map((item) => item.slot);
        }
        if (level > maxLevel) {
            maxLevel = level;
            entry = slot;
        }
    }

    function markDead(slot) {
        if (tombstones[slot] === 1) return;
        tombstones[slot] = 1;
        active -= 1;
        dead += 1;
    }

    const index = {
        get dim() { return dim; },
        get size() { return active; },
        get nodes() { return ids.length; },
        get tombstones() { return dead; },
        get m() { return m; },
        get efConstruction() { return efConstruction; },
        get efSearch() { return defaultEfSearch; },

        has(id) {
            const slot = idToSlot.get(asText(id));
            return slot !== undefined && tombstones[slot] === 0;
        },

        /** 插入或更新；返回 false 表示向量不可用（维度不对 / 零向量）。 */
        add(id, vector) {
            const key = asText(id);
            if (key === '') return false;
            const vec = normalized(vector, dim);
            if (vec === null) return false;
            if (dim === 0) dim = vec.length;
            const existing = idToSlot.get(key);
            if (existing !== undefined) {
                markDead(existing);
                idToSlot.delete(key);
            }
            insertRaw(key, vec);
            return true;
        },

        remove(id) {
            const key = asText(id);
            const slot = idToSlot.get(key);
            if (slot === undefined) return false;
            markDead(slot);
            idToSlot.delete(key);
            return true;
        },

        /** 只读检索：返回 [{id, score}]，score 是余弦（-1..1），按降序。 */
        search(vector, k, opts = {}) {
            const query = normalized(vector, dim);
            if (query === null || entry === -1 || active === 0) return [];
            const want = clampInt(k, 10, 1, 1_000);
            const filter = typeof opts.filter === 'function' ? opts.filter : null;
            let ef = clampInt(opts.efSearch, defaultEfSearch, want, ANN_MAX_EF);
            if (ef < want) ef = want;
            let current = { slot: entry, sim: dot(query, vectors[entry]) };
            for (let levelAt = maxLevel; levelAt >= 1; levelAt -= 1) current = greedy(query, current.slot, levelAt);
            let found = [];
            for (let attempt = 0; attempt < 3; attempt += 1) {
                const layer = searchLayer(query, [current.slot], ef, 0);
                found = [];
                for (const item of layer) {
                    if (tombstones[item.slot] === 1) continue;
                    const id = ids[item.slot];
                    if (filter !== null && !filter(id)) continue;
                    found.push({ id, score: item.sim });
                    if (filter === null && found.length >= want) break;
                }
                if (found.length >= want || filter === null) break;
                // 过滤后不够：加大 beam 再走一遍；到探针上限就交给调用方的精确扫描。
                const next = Math.min(ANN_MAX_EF, ef * 2);
                if (next === ef) break;
                ef = next;
            }
            return found.slice(0, want);
        },

        /** 存活 id 的只读快照（重建、测试用）。 */
        liveIds() {
            const out = [];
            for (let i = 0; i < ids.length; i += 1) if (tombstones[i] === 0) out.push(ids[i]);
            return out;
        },

        /** 墓碑过多时按存活向量整图重建（同一批 id 与向量，图重排）。 */
        compact() {
            if (dead === 0) return { rebuilt: false, active };
            const live = [];
            for (let i = 0; i < ids.length; i += 1) if (tombstones[i] === 0) live.push({ id: ids[i], vector: vectors[i] });
            reset();
            for (const item of live) insertRaw(item.id, item.vector);
            return { rebuilt: true, active };
        },

        stats() {
            let linkCount = 0;
            for (const perNode of links) for (const level of perNode) if (Array.isArray(level)) linkCount += level.length;
            return {
                dim, nodes: ids.length, active, tombstones: dead, m, efConstruction,
                maxLevel: maxLevel + 1, links: linkCount, bytes: 0,
            };
        },

        /** 序列化成单文件二进制（小端）。 */
        toBuffer() {
            const pieces = [];
            const header = Buffer.alloc(HEADER_BYTES);
            header.write(ANN_MAGIC, 0, 4, 'ascii');
            header.writeUInt32LE(ANN_VERSION, 4);
            header.writeUInt32LE(dim >>> 0, 8);
            header.writeUInt32LE(ids.length >>> 0, 12);
            header.writeUInt32LE(active >>> 0, 16);
            header.writeUInt32LE(dead >>> 0, 20);
            header.writeUInt32LE(m >>> 0, 24);
            header.writeUInt32LE(efConstruction >>> 0, 28);
            header.writeInt32LE(entry, 32);
            header.writeInt32LE(maxLevel, 36);
            header.writeUInt32LE(seed >>> 0, 40);
            pieces.push(header);
            for (let slot = 0; slot < ids.length; slot += 1) {
                const idBytes = Buffer.from(ids[slot], 'utf8');
                const levelCount = levels[slot] + 1;
                let linkBytes = 0;
                for (let levelAt = 0; levelAt < levelCount; levelAt += 1) linkBytes += 2 + (links[slot][levelAt]?.length ?? 0) * 4;
                const record = Buffer.alloc(4 + idBytes.length + dim * 4 + linkBytes);
                let at = 0;
                record.writeUInt16LE(idBytes.length, at); at += 2;
                record.writeUInt8(Math.min(255, levels[slot]), at); at += 1;
                record.writeUInt8(tombstones[slot], at); at += 1;
                idBytes.copy(record, at); at += idBytes.length;
                for (let i = 0; i < dim; i += 1) {
                    record.writeFloatLE(vectors[slot][i], at);
                    at += 4;
                }
                for (let levelAt = 0; levelAt < levelCount; levelAt += 1) {
                    const neighbors = links[slot][levelAt] ?? [];
                    record.writeUInt16LE(neighbors.length, at); at += 2;
                    for (const neighbor of neighbors) {
                        record.writeUInt32LE(neighbor >>> 0, at);
                        at += 4;
                    }
                }
                pieces.push(record);
            }
            return Buffer.concat(pieces);
        },

        reset,

        /**
         * 装回磁盘上的图（parseAnnBuffer 用）：`add()` 会按新图重建、慢且与磁盘
         * 里的连接不同，这里要的是逐字节一致的原样恢复。
         */
        __restore({ dim: loadedDim, entry: loadedEntry, maxLevel: loadedMaxLevel, nodes }) {
            reset();
            dim = loadedDim;
            for (const node of nodes) {
                const slot = grow();
                ids[slot] = node.id;
                vectors[slot] = node.vector;
                levels[slot] = node.level;
                tombstones[slot] = node.tomb === true ? 1 : 0;
                links[slot] = node.links;
                idToSlot.set(node.id, slot);
                if (node.tomb === true) dead += 1; else active += 1;
            }
            entry = loadedEntry;
            maxLevel = loadedMaxLevel;
        },
    };

    function reset() {
        ids.length = 0;
        vectors.length = 0;
        levels.length = 0;
        tombstones.length = 0;
        links.length = 0;
        idToSlot.clear();
        entry = -1;
        maxLevel = -1;
        active = 0;
        dead = 0;
    }

    if (dim > 0) { /* 维度可以等到第一条向量再定 */ }
    return index;
}

/**
 * 反序列化。任何结构问题都抛错 —— 调用方（kb-db）据此重建索引，而不是带着一张
 * 半截的图继续检索。
 */
export function parseAnnBuffer(buffer) {
    if (!Buffer.isBuffer(buffer) || buffer.length < HEADER_BYTES) throw new Error('ANN 文件太短');
    if (buffer.toString('ascii', 0, 4) !== ANN_MAGIC) throw new Error('ANN 文件魔数不对');
    const version = buffer.readUInt32LE(4);
    if (version !== ANN_VERSION) throw new Error(`ANN 文件版本 ${version} 读不懂`);
    const dim = buffer.readUInt32LE(8);
    const nodeCount = buffer.readUInt32LE(12);
    const active = buffer.readUInt32LE(16);
    const dead = buffer.readUInt32LE(20);
    const m = buffer.readUInt32LE(24);
    const efConstruction = buffer.readUInt32LE(28);
    const entry = buffer.readInt32LE(32);
    const maxLevel = buffer.readInt32LE(36);
    const seed = buffer.readUInt32LE(40);
    if (dim === 0 || dim > 4096) throw new Error('ANN 文件维度不合法');
    const index = createAnnIndex({ dim, m, efConstruction, seed });
    // createAnnIndex 的 add() 会重建自己的图（慢且与磁盘里的图不同），这里要的是
    // 「原样装回」：直接搬槽位、层、墓碑与邻居表。
    index.__restore({
        dim, entry, maxLevel,
        nodes: readNodes(buffer, nodeCount, dim),
    });
    if (index.size !== active || index.tombstones !== dead) {
        throw new Error(`ANN 文件计数对不上：存活 ${index.size}/${active}，墓碑 ${index.tombstones}/${dead}`);
    }
    return index;
}

function readNodes(buffer, nodeCount, dim) {
    const nodes = [];
    let at = HEADER_BYTES;
    for (let slot = 0; slot < nodeCount; slot += 1) {
        if (at + 4 > buffer.length) throw new Error('ANN 文件在节点头处截断');
        const idLen = buffer.readUInt16LE(at); at += 2;
        const level = buffer.readUInt8(at); at += 1;
        const tomb = buffer.readUInt8(at); at += 1;
        if (at + idLen + dim * 4 > buffer.length) throw new Error('ANN 文件在 id / 向量处截断');
        const id = buffer.toString('utf8', at, at + idLen); at += idLen;
        const vector = new Float32Array(dim);
        for (let i = 0; i < dim; i += 1) { vector[i] = buffer.readFloatLE(at); at += 4; }
        const perLevel = [];
        for (let levelAt = 0; levelAt <= level; levelAt += 1) {
            if (at + 2 > buffer.length) throw new Error('ANN 文件在邻居数处截断');
            const count = buffer.readUInt16LE(at); at += 2;
            if (at + count * 4 > buffer.length) throw new Error('ANN 文件在邻居表处截断');
            const neighbors = new Array(count);
            for (let i = 0; i < count; i += 1) { neighbors[i] = buffer.readUInt32LE(at); at += 4; }
            perLevel.push(neighbors);
        }
        nodes.push({ id, level, tomb: tomb === 1, vector, links: perLevel });
    }
    return nodes;
}

/**
 * 路径约定：`<知识库目录>/index/ann-<模型>-<维度>.bin`，旁边一份 manifest。
 * 文件名里放模型与维度，换模型天然不会读到旧图（旧文件留给清理脚本）。
 */
export function annSlugOf(model, dim) {
    const safe = asText(model).replace(/[^a-zA-Z0-9._-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 64) || 'unknown';
    return `${safe}-d${Math.max(0, Math.trunc(Number(dim) || 0))}`;
}

export function annPathsOf(dir, model, dim) {
    const slug = annSlugOf(model, dim);
    const indexDir = join(asText(dir), 'index');
    return { dir: indexDir, file: join(indexDir, `ann-${slug}.bin`), manifest: join(indexDir, `ann-${slug}.json`) };
}

/**
 * 落盘：文件与 manifest 都先写临时名再 rename。
 * 多进程同时写时后写者覆盖前者；索引是缓存，谁赢都不丢真源。
 */
export function saveAnn(index, { dir, model, updatedAt = '', count = 0 } = {}) {
    if (index.nodes === 0 || index.dim === 0) throw new Error('空索引不落盘（没有向量时不该有索引文件）');
    const paths = annPathsOf(dir, model, index.dim);
    mkdirSync(paths.dir, { recursive: true });
    const buffer = index.toBuffer();
    const tmp = `${paths.file}.tmp-${process.pid}`;
    writeFileSync(tmp, buffer);
    renameSync(tmp, paths.file);
    const manifest = {
        version: ANN_VERSION,
        model: asText(model),
        dim: index.dim,
        nodes: index.nodes,
        active: index.size,
        tombstones: index.tombstones,
        count,
        updatedAt: asText(updatedAt),
        builtAt: new Date().toISOString(),
        file: paths.file,
        bytes: buffer.length,
        m: index.m,
        efConstruction: index.efConstruction,
    };
    const tmpManifest = `${paths.manifest}.tmp-${process.pid}`;
    writeFileSync(tmpManifest, JSON.stringify(manifest, null, 2) + '\n', 'utf8');
    renameSync(tmpManifest, paths.manifest);
    return manifest;
}

/**
 * 从磁盘装回；manifest 与库里的 embeddings 计数/更新时间对不上、文件损坏、
 * 维度不一致都返回 null（调用方重建）。读盘失败不抛。
 */
export function loadAnn({ dir, model, dim, count = null, updatedAt = null } = {}) {
    const paths = annPathsOf(dir, model, dim);
    try {
        const manifest = JSON.parse(readFileSync(paths.manifest, 'utf8'));
        if (manifest === null || typeof manifest !== 'object') return null;
        if (Number(manifest.dim) !== Number(dim)) return null;
        if (asText(manifest.model) !== asText(model)) return null;
        if (count !== null && Number(manifest.count) !== Number(count)) return null;
        if (updatedAt !== null && asText(manifest.updatedAt) !== asText(updatedAt)) return null;
        const buffer = readFileSync(paths.file);
        const index = parseAnnBuffer(buffer);
        if (index.dim !== Number(dim)) return null;
        return { index, manifest };
    } catch {
        return null;
    }
}

/** 删掉一份索引（重建前 / 清理损坏文件用）。 */
export function dropAnn({ dir, model, dim } = {}) {
    const paths = annPathsOf(dir, model, dim);
    for (const target of [paths.file, paths.manifest]) {
        try { rmSync(target, { force: true }); } catch { /* 删不掉不影响重建 */ }
    }
}

function clampInt(value, fallback, min, max) {
    const n = typeof value === 'number' ? value : Number.parseFloat(String(value));
    if (!Number.isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, Math.trunc(n)));
}
