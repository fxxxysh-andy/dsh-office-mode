/**
 * 最小 XML 工具：转义、解析、遍历。
 *
 * 只覆盖 OOXML 实际出现的语法：声明、注释、CDATA、DOCTYPE、自闭合标签、
 * 单双引号属性、五个预定义实体和数字实体。刻意不做命名空间解析 —— OOXML
 * 的标签都带前缀（`w:p`、`x:row`），按前缀原名比较反而更直接。
 *
 * @module dsh-office-mode/engine/xml
 */

const XML_ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' };

/** 转义文本节点内容（属性值同样适用）。 */
export function escapeXml(value) {
    return String(value ?? '').replace(/[&<>"']/g, (ch) => XML_ESCAPES[ch]);
}

/** 去掉标签前缀：`w:p` → `p`。 */
export function localName(name) {
    const at = name.indexOf(':');
    return at === -1 ? name : name.slice(at + 1);
}

function decodeEntities(text) {
    if (text.indexOf('&') === -1) return text;
    return text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, body) => {
        if (body[0] === '#') {
            const hex = body[1] === 'x' || body[1] === 'X';
            const code = Number.parseInt(hex ? body.slice(2) : body.slice(1), hex ? 16 : 10);
            return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
        }
        switch (body) {
            case 'amp': return '&';
            case 'lt': return '<';
            case 'gt': return '>';
            case 'quot': return '"';
            case 'apos': return "'";
            default: return whole;
        }
    });
}

function parseAttrs(source) {
    const attrs = {};
    const re = /([^\s=/]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
    let match;
    while ((match = re.exec(source)) !== null) {
        attrs[match[1]] = decodeEntities(match[3] !== undefined ? match[3] : match[4]);
    }
    return attrs;
}

/**
 * 解析 XML 文本，返回根元素。
 * 元素节点形状：`{ name, attrs, children, text }`；`text` 是直接子文本节点的拼接。
 */
export function parseXml(source) {
    const root = { name: '#document', attrs: {}, children: [], text: '' };
    const stack = [root];
    const len = source.length;
    let at = 0;
    while (at < len) {
        const lt = source.indexOf('<', at);
        if (lt === -1) {
            appendText(stack[stack.length - 1], source.slice(at));
            break;
        }
        if (lt > at) appendText(stack[stack.length - 1], source.slice(at, lt));
        if (source.startsWith('<!--', lt)) {
            const end = source.indexOf('-->', lt + 4);
            at = end === -1 ? len : end + 3;
            continue;
        }
        if (source.startsWith('<![CDATA[', lt)) {
            const end = source.indexOf(']]>', lt + 9);
            const body = source.slice(lt + 9, end === -1 ? len : end);
            stack[stack.length - 1].text += body;
            at = end === -1 ? len : end + 3;
            continue;
        }
        if (source.startsWith('<?', lt) || source.startsWith('<!', lt)) {
            const end = source.indexOf('>', lt);
            at = end === -1 ? len : end + 1;
            continue;
        }
        const gt = findTagEnd(source, lt);
        if (gt === -1) {
            appendText(stack[stack.length - 1], source.slice(lt));
            break;
        }
        const inner = source.slice(lt + 1, gt);
        if (inner[0] === '/') {
            const name = inner.slice(1).trim();
            if (stack.length > 1) {
                const top = stack[stack.length - 1];
                if (top.name === name) stack.pop();
                else {
                    // 容错：往上找同名祖先，找不到就忽略这个闭合标签
                    for (let i = stack.length - 1; i > 0; i -= 1) {
                        if (stack[i].name === name) {
                            stack.length = i;
                            break;
                        }
                    }
                }
            }
        } else {
            const selfClosing = inner.endsWith('/');
            const body = selfClosing ? inner.slice(0, -1) : inner;
            const space = body.search(/[\s/]/);
            const name = space === -1 ? body : body.slice(0, space);
            const node = {
                name,
                attrs: parseAttrs(space === -1 ? '' : body.slice(space)),
                children: [],
                text: '',
            };
            stack[stack.length - 1].children.push(node);
            if (!selfClosing) stack.push(node);
        }
        at = gt + 1;
    }
    return root;
}

function appendText(node, raw) {
    if (node === undefined) return;
    const text = decodeEntities(raw);
    if (text !== '') node.text += text;
}

/** 找到标签真正结束的 `>`，跳过引号里的内容。 */
function findTagEnd(source, from) {
    let quote = '';
    for (let i = from + 1; i < source.length; i += 1) {
        const ch = source[i];
        if (quote !== '') {
            if (ch === quote) quote = '';
        } else if (ch === '"' || ch === "'") {
            quote = ch;
        } else if (ch === '>') {
            return i;
        }
    }
    return -1;
}

/** 直接子元素，可按名字过滤（比较带前缀的原名）。 */
export function children(node, name) {
    if (!node) return [];
    if (name === undefined) return node.children;
    return node.children.filter((child) => child.name === name);
}

/** 深度优先查找全部后代。 */
export function descendants(node, name) {
    const out = [];
    const walk = (current) => {
        for (const child of current.children) {
            if (name === undefined || child.name === name) out.push(child);
            walk(child);
        }
    };
    if (node) walk(node);
    return out;
}

/** 第一个匹配的后代。 */
export function descendant(node, name) {
    return descendants(node, name)[0];
}

/** 读属性。 */
export function attr(node, name, fallback = undefined) {
    const value = node?.attrs?.[name];
    return value === undefined ? fallback : value;
}

/** 递归拼接元素内的全部文本。 */
export function textOf(node) {
    if (!node) return '';
    let out = node.text;
    for (const child of node.children) out += textOf(child);
    return out;
}

/** 解析 `xml:space` 之外无意义的空白折叠（用于对比标题文本）。 */
export function normalizeSpace(value) {
    return String(value ?? '').replace(/\s+/g, ' ').trim();
}
