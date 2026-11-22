/**
 * DeepSeek 官方 OCR 通道（第九十一轮，2026-11-19）。
 *
 * 契约来源（逐字核对过官方文档，不是按记忆写的）：
 *   - API 参考：https://api-docs.deepseek.com/zh-cn/api/ocr.md
 *   - 指南：    https://api-docs.deepseek.com/zh-cn/guides/ocr.md
 *
 * 官方契约摘要：
 *   1. 端点：POST https://api.deepseek.com/v1/tools/ocr（Bearer 鉴权）。
 *   2. 请求体是 multipart/form-data，只有两个字段：
 *      - file：要识别的文件，必须带文件名；
 *      - purpose：text_recognition | rich_layout_recognition。
 *      多传别的字段官方明确返回 400 —— 所以这里只发这两个。
 *   3. 支持 PDF / Word(DOC,DOCX) / PowerPoint(PPT,PPTX) / png jpeg webp gif bmp tiff；
 *      格式按文件实际内容判断。上限 100 MiB、1000 页；上传 120s 内完成、总时长 180s。
 *   4. 响应：识别过程中服务端持续发换行保持连接，最后是一个 JSON 对象
 *      （换行是 JSON 的前导空白，整包可以直接 JSON.parse）：
 *      { id, object:"ocr.result", created, purpose, page_count,
 *        pages:[{index,text,finish_reason}], usage }。
 *   5. 识别开始后 HTTP 状态已经是 200：失败时响应体不是识别结果，而是
 *      { error: { message, type, param, code } }。所以 200 也要查 error。
 *   6. 错误码：400 表单/purpose 不合法或文件无法处理；402 余额不足；408 上传过慢；
 *      413 文件过大；429 并发超限（每账号最多 15 个并发识别）；500/503/504 服务端问题。
 *   7. 计费：每 1000 页 1.5 元，按成功页数计；失败/超时不计费。
 *
 * 本模块只做「调用 + 解析 + 把失败翻译成人话」，不碰知识库与文件选择。
 *
 * @module dsh-office-mode/ocr
 */

/** 官方允许的 purpose（枚举只有这两个）。 */
export const OCR_PURPOSES = ['text_recognition', 'rich_layout_recognition'];

/** 官方文件体积上限：100 MiB。 */
export const OCR_MAX_BYTES = 100 * 1024 * 1024;

/** 官方页数上限：1000 页。 */
export const OCR_MAX_PAGES = 1000;

/** 官方 OCR 的 origin；复用了宿主官方供应商 Key 的请求只能发到这里。 */
export const OFFICIAL_OCR_ORIGIN = 'https://api.deepseek.com';

/** 官方 OCR 的精确端点；不复用任何可配置 baseURL。 */
export const OFFICIAL_OCR_ENDPOINT = 'https://api.deepseek.com/v1/tools/ocr';

/** 官方接受的扩展名（格式其实由内容判断，这里只用来给前端早退与提示）。 */
export const OCR_SUPPORTED_EXTENSIONS = [
    '.pdf', '.doc', '.docx', '.ppt', '.pptx',
    '.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.tif', '.tiff',
];

/** 默认配置：全空也能用（enabled=true 但没 Key 时给明确提示，不是静默失败）。 */
export const DEFAULT_OCR = {
    enabled: true,
    /** DeepSeek 开放平台的根地址；末尾斜杠会被去掉，拼 /v1/tools/ocr 或 /tools/ocr。 */
    baseURL: 'https://api.deepseek.com',
    apiKey: '',
    apiKeyEnv: 'DEEPSEEK_API_KEY',
    purpose: 'text_recognition',
    /** 与官方总时长 180s 对齐；上传阶段单独按 120s 处理由服务端决定。 */
    timeoutMs: 180_000,
};

/** 把 OCR 的失败翻译成一句能照做的话，并标出「重试有没有用」。 */
const OCR_STATUS_EXPLAIN = {
    400: { message: '表单或 purpose 不合法，或者文件无法处理（加密、格式不支持、图片解码失败）', retryable: false },
    402: { message: 'DeepSeek 账户余额不足', retryable: false },
    408: { message: '上传太慢：需在 120 秒内传完，且中途停顿不超过 30 秒', retryable: true },
    413: { message: '文件超过 100 MiB', retryable: false },
    429: { message: '同时进行的识别超过上限（每账号最多 15 个），稍后再试', retryable: true },
    500: { message: '服务端文件上传或提交识别失败', retryable: true },
    503: { message: '服务繁忙', retryable: true },
    504: { message: '上传与提交超过总时长（180 秒）', retryable: true },
};

/** OCR 通道的错误：带 status / code / retryable，面板上可以直接分层显示。 */
export class OcrError extends Error {
    /**
     * @param {string} message 给人看的一句话（已经包含官方返回的 message）
     * @param {{status?: number, code?: string, retryable?: boolean, detail?: string, hint?: string,
     *          contentType?: string, bodyKind?: string}} [options]
     *   contentType / bodyKind 是「上游响应体的公开分类」（空 / HTML / 非 JSON 文本 / JSON），
     *   给端点原样回给页面做诊断用 —— 不带响应体原文。
     */
    constructor(message, options = {}) {
        super(message);
        this.name = 'OcrError';
        this.status = options.status ?? 0;
        this.code = options.code ?? '';
        this.retryable = options.retryable === true;
        this.detail = options.detail ?? '';
        this.hint = options.hint ?? '';
        this.contentType = options.contentType ?? '';
        this.bodyKind = options.bodyKind ?? '';
    }
}

/**
 * 从设置里那份原始配置解析出 OCR 配置（不填就是这个模块的默认值）。
 * @param {object} [raw] ocr 那一段原始配置
 */
export function resolveOcrConfig(raw) {
    const source = raw && typeof raw === 'object' ? raw : {};
    const purpose = String(source.purpose ?? '').trim();
    const timeoutMs = Number(source.timeoutMs);
    return {
        enabled: source.enabled !== false,
        baseURL: String(source.baseURL ?? '').trim() || DEFAULT_OCR.baseURL,
        apiKey: String(source.apiKey ?? '').trim(),
        apiKeyEnv: String(source.apiKeyEnv ?? '').trim(),
        purpose: OCR_PURPOSES.includes(purpose) ? purpose : DEFAULT_OCR.purpose,
        timeoutMs: Number.isFinite(timeoutMs) && timeoutMs > 0 ? timeoutMs : DEFAULT_OCR.timeoutMs,
    };
}

/**
 * 把配置里的环境变量名收敛成一个能安全写进错误文案的标签。
 *
 * 为什么不能直接拼：`ocr.apiKeyEnv` 只是设置页里的一个普通文本框，用户很容易把 Key
 * 本身粘进去（粘贴比手敲变量名自然）。真那样时，把这段字符串原样拼进错误消息，就等于
 * 把密钥送回了浏览器 —— 所以这里只认「像环境变量名」的取值（字母或下划线开头，其余
 * 是字母 / 数字 / 下划线，最长 128），不合法就返回空串，由调用方换成不带名字的说法。
 * 这条规则**只影响文案**：读不读得到 Key 仍由调用方按配置里的原始名字去查。
 *
 * 空值不算「不合法」：没填时按默认变量名说，保持旧文案不变。
 *
 * @param {unknown} raw 设置里的 ocr.apiKeyEnv（只来自配置，绝不来自请求体）
 * @returns {string} 可以显示的变量名，或 ''（不能安全显示）
 */
export function safeOcrKeyEnvLabel(raw) {
    const name = String(raw ?? '').trim();
    if (name === '') return DEFAULT_OCR.apiKeyEnv;
    return /^[A-Za-z_][A-Za-z0-9_]{0,127}$/.test(name) ? name : '';
}

/**
 * 拼端点。两种写法都收：根地址（https://api.deepseek.com）与已经带 /v1 的地址。
 * @param {string} baseURL
 */
export function ocrEndpoint(baseURL) {
    const base = String(baseURL ?? '').trim().replace(/\/+$/, '');
    if (base === '') return OFFICIAL_OCR_ENDPOINT;
    return /\/v1$/i.test(base) ? base + '/tools/ocr' : base + '/v1/tools/ocr';
}

/**
 * 把设置里的 baseURL 收敛成「精确官方端点」，或返回 undefined。
 *
 * 复用了宿主官方供应商 Key 时，只在目的地**就是**官方 OCR 时继续：
 *   - 协议必须是 https；
 *   - 不能带 userinfo（user:pass@）、端口（443 由 URL 归一化后为空）、
 *     query 或 fragment；
 *   - hostname 必须精确等于 api.deepseek.com；
 *   - path 只能是空 / / / /v1 / /v1/（插件自己补 /tools/ocr）。
 *
 * 这不是「把自定义地址改写掉」：调用方在返回 undefined 时必须放弃复用供应商
 * Key，而不是偷偷换成官方地址 —— 用户把 OCR 配到别处，内容就只该去别处。
 *
 * @param {unknown} baseURL 设置里的 ocr.baseURL（空 = 默认官方）
 * @returns {string | undefined} 官方精确端点，或 undefined（不得复用）
 */
export function officialOcrEndpointOf(baseURL) {
    const raw = String(baseURL ?? '').trim();
    if (raw === '') return OFFICIAL_OCR_ENDPOINT;
    let url;
    try {
        url = new URL(raw);
    } catch {
        return undefined;
    }
    if (url.protocol !== 'https:') return undefined;
    if (url.username !== '' || url.password !== '') return undefined;
    if (url.port !== '' && url.port !== '443') return undefined;
    if (url.search !== '' || url.hash !== '') return undefined;
    if (url.hostname !== 'api.deepseek.com') return undefined;
    if (url.pathname !== '' && url.pathname !== '/' && url.pathname !== '/v1' && url.pathname !== '/v1/') {
        return undefined;
    }
    return OFFICIAL_OCR_ENDPOINT;
}

/**
 * 把官方**结构化**的错误对象翻成一句话。
 *
 * 只认 body.error 对象：纯文本/HTML 响应体一律不引用原文，交给
 * describeUnusableBody 做公开分类 —— 中间隔着代理时，那段原文可能是网关错误页、
 * 内网报文或带本机路径的文本，拼进消息再随 JSON 回到浏览器既是误导也是外泄。
 */
function describeBodyError(body) {
    const error = body && typeof body === 'object' ? body.error : undefined;
    if (error && typeof error === 'object') {
        const code = String(error.code ?? error.type ?? '').trim();
        const message = String(error.message ?? '').trim() || '没有给出原因';
        return { code, message };
    }
    return { code: '', message: '' };
}

/**
 * 上游响应体只做公开分类，不保留原文。
 *
 * 200 也可能是失败（识别开始后状态码已固定成 200，失败装在 body.error 里），
 * 而中间隔着反向代理时，200/4xx/5xx 的响应体可能是网关的 HTML 错误页或内网报文。
 * 这里只回报状态码、content-type 与「空 / HTML-XML / 非 JSON 文本 / JSON」分类，
 * 原文一个字都不回显。
 *
 * @param {{status?: number, headers?: {get?: Function}}} response
 * @param {unknown} body 已解析的响应体（不是 JSON 时为字符串，空响应为 null）
 * @param {string} rawText 原始文本；只用来判空与 HTML 前缀，不参与输出
 * @returns {{kind: string, label: string, contentType: string, text: string}}
 */
function describeUnusableBody(response, body, rawText) {
    const getHeader = response && response.headers && typeof response.headers.get === 'function'
        ? (name) => response.headers.get(name)
        : () => '';
    const contentType = String(getHeader('content-type') ?? '').split(';')[0].trim().toLowerCase();
    const text = typeof rawText === 'string' ? rawText.trim() : '';
    const isObject = body !== null && body !== undefined && typeof body === 'object';
    let kind;
    if (text === '') kind = 'empty';
    else if (/html|xml/.test(contentType) || /^<(?:!doctype|html|head|body|xml|meta|title|div|p)\b/i.test(text)) kind = 'html';
    else if (isObject) kind = 'json';
    else kind = 'text';
    const label = { empty: '空响应体', html: 'HTML/XML 页面', text: '非 JSON 文本', json: 'JSON' }[kind];
    const status = Number(response && response.status);
    const parts = ['HTTP ' + (Number.isInteger(status) ? status : 0)];
    parts.push(contentType === '' ? '没有 content-type' : 'content-type: ' + contentType);
    parts.push('响应体分类：' + label);
    return { kind, label, contentType: contentType, text: '（' + parts.join('，') + '）' };
}

/**
 * 只挑可安全外发的传输错误码（ECONNRESET / UND_ERR_SOCKET 这类全大写 token）。
 * 异常 message 可能带主机名、端口或本机路径，一律不拼进错误消息。
 */
function safeTransportCode(error) {
    const raw = String(error?.code ?? error?.cause?.code ?? '').trim();
    return /^[A-Z][A-Z0-9_]{2,40}$/.test(raw) ? raw : '';
}

/**
 * 把借来的 Key 从任何要外发的文本里剥掉。
 *
 * 上游错误对象、fetch 异常消息都可能原样回显 Authorization（有的网关会把
 * 「invalid key sk-...」写进 message / code）。这里只替换**这一份精确的 Key**，
 * 不做泛化的密钥扫描；空 Key 原样返回，避免改变显式 Key 的既有文案。
 *
 * @param {unknown} value
 * @param {string} secret 本次请求实际使用的 Key（只有官方供应商来源才传入）
 * @returns {string}
 */
const REDACTED_KEY = '[已隐去]';

function redactBorrowedKey(value, secret) {
    const text = String(value ?? '');
    if (secret === '') return text;
    return text.split(secret).join(REDACTED_KEY);
}

/**
 * 识别一个本地文件。
 *
 * @param {object} options
 * @param {string} options.filePath 本地绝对路径
 * @param {string} [options.purpose] text_recognition | rich_layout_recognition
 * @param {string} [options.baseURL]
 * @param {string} [options.apiKey]
 * @param {string} [options.apiKeyEnv] 设置里的环境变量名；只用来把「缺 Key」说清楚
 * @param {'official-provider'} [options.apiKeySource] Key 来自宿主官方供应商时置为
 *   official-provider：端点被钉死为官方地址，且拒绝任何重定向
 * @param {boolean} [options.hostProviderHint] 官方供应商行存在、只差 Key 时为 true，
 *   缺 Key 文案才会补一句 Models 设置
 * @param {'disabled'|'custom_endpoint'|'nonlocal'} [options.reuseHint] 这次没有借宿主
 *   Key 的公开原因：设置没打开 / 端点不是官方 / 未通过本机来源检查
 * @param {number} [options.timeoutMs]
 * @param {Function} [options.fetchImpl] 测试时注入
 * @returns {Promise<object>} { ok, id, purpose, pageCount, pages, usage, bytes, fileName, ms, requestURL }
 */
export async function ocrDocument(options = {}) {
    const filePath = String(options.filePath ?? '');
    if (filePath === '') throw new OcrError('OCR 需要一个本地文件路径。', { retryable: false });
    const purpose = OCR_PURPOSES.includes(String(options.purpose ?? ''))
        ? String(options.purpose)
        : DEFAULT_OCR.purpose;
    const baseURL = String(options.baseURL ?? '').trim() || DEFAULT_OCR.baseURL;
    const apiKey = String(options.apiKey ?? '').trim();
    const apiKeySource = String(options.apiKeySource ?? '');
    // 只有「借来的」Key 需要额外清洗：显式 Key 的既有错误文案保持不变。
    const borrowedKey = apiKeySource === 'official-provider' ? apiKey : '';
    const redact = (value) => redactBorrowedKey(value, borrowedKey);
    const hostProviderHint = options.hostProviderHint === true;
    const timeoutMs = Number.isFinite(Number(options.timeoutMs)) && Number(options.timeoutMs) > 0
        ? Number(options.timeoutMs)
        : DEFAULT_OCR.timeoutMs;
    const fetchImpl = typeof options.fetchImpl === 'function' ? options.fetchImpl : globalThis.fetch;
    if (typeof fetchImpl !== 'function') {
        throw new OcrError('当前 Node 运行时没有 fetch，OCR 通道不可用。', { retryable: false });
    }
    if (apiKey === '') {
        // 环境变量名只信设置里那一份（由 view.js 从配置透传）：请求体里的任何字段都不参与，
        // 客户端字符串一旦拼进诊断文案，就既是注入面也是外泄面。文案里也永远不出现 Key 的值。
        const envLabel = safeOcrKeyEnvLabel(options.apiKeyEnv);
        // 两句提示都指向同一件能照做的事：在设置里填，或按变量名在启动 DSH 的那个终端里配。
        const where = '在 设置 → 办公模式 → 文件识别（OCR）里填 Key';
        const how = '这个变量要在启动 DSH 的那个终端里设置（它来自那个进程继承的环境变量，'
            + '换一个终端就等于换一份），改完重启 DSH 才生效。';
        // 只有端点与官方供应商行都匹配（hostProviderHint=true）才提 Models：
        // 自定义端点下绝不能让用户以为「在 Models 里存一份就会发给这里」。
        const models = hostProviderHint
            ? '如果宿主 设置 → 模型（Models）里有 DeepSeek 官方供应商，也可以在那里保存 Key，OCR 会直接复用那一份。'
            : '';
        const reuse = options.reuseHint === 'disabled'
            ? '要复用 设置 → 模型（Models）里 DeepSeek 官方供应商的 Key，请先在 设置 → 办公模式 → 文件识别（OCR）打开「复用官方供应商 Key」；它只对通过本机来源检查的浏览器请求生效。'
            : options.reuseHint === 'custom_endpoint'
                ? '当前 OCR 地址不是官方地址（https://api.deepseek.com），不能复用官方供应商 Key；请直接填 Key，或把地址改回官方。'
                : options.reuseHint === 'nonlocal'
                    ? '复用官方供应商 Key 只对通过本机来源检查的浏览器请求生效；这次请求未通过本机来源检查，所以没有使用它。这项检查不是身份认证。'
                    : '';
        throw new OcrError(
            envLabel === ''
                // 不合法就不回显：那种值更可能是误粘进去的 Key，而不是变量名。
                ? '没有 DeepSeek API Key：设置里 ocr.apiKey 留空，ocr.apiKeyEnv 里填的环境变量名不是合法变量名（这里不回显它），也没有取到值。'
                : '没有 DeepSeek API Key：设置里 ocr.apiKey 留空，环境变量 ' + envLabel + ' 也没有值。',
            {
                code: 'missing_api_key', retryable: false,
                hint: (envLabel === ''
                    ? where + '，并把「Key 的环境变量名」改成合法变量名（例如 ' + DEFAULT_OCR.apiKeyEnv + '）；' + how
                    : where + '，或配好 ' + envLabel + '；' + how) + models + reuse,
            },
        );
    }
    // 复用供应商 Key 时，目的地必须是官方端点。view.js 在调用前已经检查过一次；
    // 这里再钉一次，防止别的调用方把自定义 baseURL 和这份 Key 拼在一起。
    if (apiKeySource === 'official-provider' && officialOcrEndpointOf(baseURL) !== OFFICIAL_OCR_ENDPOINT) {
        throw new OcrError(
            '这份 Key 来自宿主官方供应商，只发给官方 OCR 地址；当前设置里的 OCR 地址不是官方地址，因此没有发送。',
            { code: 'endpoint_not_official', retryable: false },
        );
    }

    const fs = await import('node:fs/promises');
    const path = await import('node:path');
    let stat;
    try {
        stat = await fs.stat(filePath);
    } catch (error) {
        throw new OcrError('读不到文件「' + filePath + '」：' + (error?.message ?? error), { code: 'file_missing', retryable: false });
    }
    if (!stat.isFile()) throw new OcrError('「' + filePath + '」不是文件。', { code: 'not_a_file', retryable: false });
    if (stat.size > OCR_MAX_BYTES) {
        throw new OcrError('文件有 ' + stat.size + ' 字节，超过官方 100 MiB 上限。', { code: 'file_too_large', retryable: false });
    }
    if (stat.size === 0) throw new OcrError('文件是空的。', { code: 'file_empty', retryable: false });

    const fileName = path.basename(filePath);
    // 只发官方认的两个字段（file / purpose）：多传字段官方直接 400。
    const form = new FormData();
    form.append('file', new Blob([await fs.readFile(filePath)]), fileName);
    form.append('purpose', purpose);

    // 复用供应商 Key 时用写死的官方端点；自定义配置永远不参与这条路径。
    const officialProviderKey = apiKeySource === 'official-provider';
    const requestURL = officialProviderKey ? OFFICIAL_OCR_ENDPOINT : ocrEndpoint(baseURL);
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const startedAt = Date.now();
    let response;
    try {
        const requestInit = {
            method: 'POST',
            headers: { authorization: 'Bearer ' + apiKey, accept: 'application/json' },
            body: form,
            signal: controller.signal,
        };
        // 官方端点不做任何重定向：3xx 直接失败，避免 Key 跟着 Location 走到别处。
        if (officialProviderKey) requestInit.redirect = 'error';
        response = await fetchImpl(requestURL, requestInit);
        // 双保险：即便注入的 fetch 忽略了 redirect:'error'，只要它承认跟过跳转
        // 或最终 URL 不是官方端点，就当场失败 —— 这份 Key 不能跟着 Location 走。
        if (officialProviderKey
            && (response?.redirected === true
                || (typeof response?.url === 'string' && response.url !== '' && response.url !== OFFICIAL_OCR_ENDPOINT))) {
            clearTimeout(timer);
            throw new OcrError(
                '官方 OCR 地址返回了重定向；这份 Key 只发给官方端点，没有继续。',
                { code: 'redirect_blocked', retryable: false },
            );
        }
    } catch (error) {
        clearTimeout(timer);
        // 上面主动抛的 redirect_blocked 不能被改写成网络错误。
        if (error instanceof OcrError) throw error;
        const aborted = controller.signal.aborted;
        // 借来的 Key 不能跟着任意异常消息回显（fetch 失败信息可能带上 Authorization）。
        // 只保留可执行的传输错误码；没有安全码就只说「网络错误」。
        const rawTransport = safeTransportCode(error);
        const transport = rawTransport === borrowedKey ? '' : redact(rawTransport);
        throw new OcrError(
            aborted
                ? 'OCR 请求超过 ' + timeoutMs + ' 毫秒没有返回（官方总时长上限 180 秒）。'
                : borrowedKey === ''
                    ? 'OCR 请求发不出去：' + (error?.message ?? error)
                    : 'OCR 请求发不出去：网络错误' + (transport === '' ? '' : '（' + transport + '）') + '。',
            { code: aborted ? 'timeout' : 'network', retryable: true },
        );
    }

    // 计时器**不**在这里清：官方在识别过程中持续发换行保活（心跳），headers 到了
    // body 还可能在流。超时必须从「发出请求」一直管到「响应体读完」，否则一个卡住的
    // body 会一直挂着 —— 那就是「有超时」的假象，而 abort 只发生在 headers 之前。
    let rawText = '';
    let bodyReadError = null;
    try {
        rawText = await response.text();
    } catch (error) {
        bodyReadError = error;
    } finally {
        clearTimeout(timer);
    }
    // 耗时按「整份响应收完」计：只算到 headers 的时间不是用户感知的识别耗时。
    const ms = Date.now() - startedAt;
    if (bodyReadError !== null) {
        // 读 body 失败有两种，绝不能混成「空响应」：
        //   - 超时：timer 触发 abort → timeout（端点映射 504）；
        //   - 连接中途断了：body_interrupted（端点映射 502），保留上游 HTTP 状态做诊断。
        // 异常文字不外发；只附带 ECONNRESET 这类安全的传输错误码。
        const aborted = controller.signal.aborted;
        const transport = safeTransportCode(bodyReadError);
        throw new OcrError(
            aborted
                ? 'OCR 响应体超过 ' + timeoutMs + ' 毫秒没有读完（上游 HTTP ' + response.status + '，官方总时长上限 180 秒）。'
                : 'OCR 响应体没有读完就断开了（上游 HTTP ' + response.status
                    + (transport === '' ? '' : '；连接错误：' + transport) + '）。',
            {
                status: response.status,
                code: aborted ? 'timeout' : 'body_interrupted',
                retryable: true,
            },
        );
    }
    let body = null;
    // 官方响应前面可能有一串换行（心跳）；换行属于 JSON 前导空白，trim 后直接 parse。
    const trimmed = rawText.trim();
    if (trimmed !== '') {
        try {
            body = JSON.parse(trimmed);
        } catch {
            body = trimmed;
        }
    }

    if (!response.ok) {
        const explained = OCR_STATUS_EXPLAIN[response.status] ?? { message: 'HTTP ' + response.status, retryable: response.status >= 500 };
        // 官方结构化 body.error 照旧引用（那是契约里的字段，页面需要它）；
        // 纯文本/HTML 响应体只做分类，绝不回显原文。
        const structured = body && typeof body === 'object' && body.error && typeof body.error === 'object';
        const detail = structured ? describeBodyError(body) : { code: '', message: '' };
        // 上游可能把 Key 回显在 message / code 里；投影前先剥掉这一份精确的 Key。
        detail.code = redact(detail.code);
        detail.message = redact(detail.message);
        const unusable = structured ? null : describeUnusableBody(response, body, rawText);
        const suffix = structured
            ? (detail.message === '' ? '' : ' —— ' + detail.message)
            : '；响应体不是可引用的 JSON：' + unusable.label
                + (unusable.contentType === '' ? '' : '（content-type: ' + unusable.contentType + '）');
        // 如果上游 code 整个就是那把 Key，剥掉后为空：改用可执行的 HTTP 状态码。
        const projectedCode = detail.code === REDACTED_KEY ? '' : detail.code;
        throw new OcrError('OCR 失败（HTTP ' + response.status + '）：' + explained.message + suffix, {
            status: response.status,
            code: projectedCode || 'http_' + response.status,
            retryable: explained.retryable,
            detail: detail.message,
            contentType: unusable === null ? '' : unusable.contentType,
            bodyKind: unusable === null ? '' : unusable.kind,
        });
    }

    // 200 也可能是失败：识别开始后状态码已经是 200，失败装在 error 里。
    if (body === null || typeof body !== 'object') {
        // 只回报「状态 + content-type + 空/HTML/非 JSON 文本」；上游原文（如网关 HTML
        // 或带本机路径的报错）不拼进消息，避免它随端点的 JSON 一起回到浏览器。
        const unusable = describeUnusableBody(response, body, rawText);
        throw new OcrError('OCR 返回的不是识别结果 JSON' + unusable.text + '。', {
            status: response.status,
            code: 'bad_response',
            retryable: true,
            contentType: unusable.contentType,
            bodyKind: unusable.kind,
            hint: '核对 ocr.baseURL 是否指向官方 /v1 端点，以及中间是否有代理/网关拦截了请求。',
        });
    }
    if (body.error !== undefined && body.error !== null) {
        const detail = describeBodyError(body);
        // 200 也可能失败，且同样可能把 Key 回显在结构化 error 里。
        detail.code = redact(detail.code);
        detail.message = redact(detail.message);
        const transient = /temporarily unavailable|Internal Server Error|did not finish/i.test(detail.message);
        const projectedCode = detail.code === REDACTED_KEY ? '' : detail.code;
        throw new OcrError('OCR 识别失败：' + detail.message, {
            status: response.status,
            code: projectedCode || 'ocr_error',
            retryable: transient,
            detail: detail.message,
        });
    }
    if (!Array.isArray(body.pages)) {
        throw new OcrError('OCR 返回里没有 pages 数组（契约是 pages:[{index,text,finish_reason}]）。', { code: 'bad_response', retryable: false });
    }
    return {
        ok: true,
        id: typeof body.id === 'string' ? body.id : '',
        purpose: typeof body.purpose === 'string' ? body.purpose : purpose,
        pageCount: Number.isFinite(Number(body.page_count)) ? Number(body.page_count) : body.pages.length,
        pages: body.pages.map((page) => ({
            index: Number.isFinite(Number(page?.index)) ? Number(page.index) : 0,
            text: typeof page?.text === 'string' ? page.text : '',
            finishReason: typeof page?.finish_reason === 'string' ? page.finish_reason : 'stop',
        })),
        usage: body.usage && typeof body.usage === 'object' ? body.usage : {},
        bytes: stat.size,
        fileName,
        ms,
        requestURL,
    };
}

/**
 * 去掉逐页识别里的版面标记（ref / box 标签与归一化坐标），只留文字。
 * 入库/检索要的是文字，坐标只有前端裁图才用得到。
 * @param {string} text
 */
export function cleanOcrText(text) {
    return String(text ?? '')
        .replace(/<｜?\/?(?:ref|box)｜?>/g, '')
        .replace(/\[\[\s*-?\d+(?:\.\d+)?(?:\s*,\s*-?\d+(?:\.\d+)?){3}\s*\]\]/g, '')
        .replace(/[ \t]+\n/g, '\n')
        .trim();
}

/**
 * 把一次识别结果拼成一段纯文本（按页分隔），给知识库入库用。
 * @param {{pages: Array<{index: number, text: string, finishReason: string}>}} result
 * @param {{stripMarkup?: boolean, perPage?: boolean}} [options]
 */
export function ocrResultText(result, options = {}) {
    const strip = options.stripMarkup !== false;
    const pages = Array.isArray(result?.pages) ? result.pages : [];
    return pages
        .map((page) => {
            const text = strip ? cleanOcrText(page.text) : String(page.text ?? '');
            return text;
        })
        .filter((text) => text !== '')
        .join(options.perPage === false ? '\n' : '\n\n');
}
