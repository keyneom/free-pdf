/**
 * Email Templates - Store, CRUD, placeholders, import/export
 * Templates are saved in localStorage. Export/import JSON for sync across devices.
 */

const STORAGE_KEY = 'free-pdf-email-templates';
const PLACEHOLDERS = [
    '{{filename}}',
    '{{date}}',
    '{{signatureSummary}}',
    '{{signerNames}}',
    '{{pageCount}}',
    '{{documentHash}}',
    '{{attachmentNote}}'
];

const DEFAULT_TEMPLATE = {
    id: 'default',
    name: 'Default',
    subject: '{{filename}}',
    body: `Please find attached {{filename}}.

This document may contain electronic signatures; where present, signer identity, date, and document association are recorded.

Document summary:
- Pages: {{pageCount}}
{{signatureSummary}}
{{documentHash}}

Please retain this message and the attached file for your records.

{{attachmentNote}}`,
    isDefault: true,
    builtin: true
};

function loadStore() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return { version: 1, defaultId: 'default', templates: [DEFAULT_TEMPLATE] };
        const data = JSON.parse(raw);
        if (!Array.isArray(data.templates)) data.templates = [DEFAULT_TEMPLATE];
        if (!data.templates.some((t) => t.id === 'default')) {
            data.templates.unshift({ ...DEFAULT_TEMPLATE });
        }
        data.version = data.version || 1;
        data.defaultId = data.defaultId || 'default';
        return data;
    } catch {
        return { version: 1, defaultId: 'default', templates: [DEFAULT_TEMPLATE] };
    }
}

function saveStore(store) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
}

function uid() {
    return 'tpl-' + Math.random().toString(36).slice(2, 12);
}

/**
 * @typedef {Object} EmailTemplate
 * @property {string} id
 * @property {string} name
 * @property {string} subject
 * @property {string} body
 * @property {boolean} [isDefault]
 * @property {boolean} [builtin]
 */

export const emailTemplates = {
    PLACEHOLDERS,

    getTemplates() {
        const s = loadStore();
        return [...s.templates];
    },

    getDefault() {
        const s = loadStore();
        const t = s.templates.find((x) => x.id === (s.defaultId || 'default'));
        return t || s.templates[0] || DEFAULT_TEMPLATE;
    },

    getById(id) {
        return loadStore().templates.find((t) => t.id === id) || null;
    },

    /**
     * @param {Omit<EmailTemplate, 'id'>} t
     * @returns {EmailTemplate}
     */
    add(t) {
        const store = loadStore();
        const id = uid();
        const next = { id, name: t.name || 'Untitled', subject: t.subject || '', body: t.body || '', isDefault: false };
        store.templates.push(next);
        if (store.templates.length === 1) store.defaultId = id;
        saveStore(store);
        return next;
    },

    /**
     * @param {string} id
     * @param {Partial<Pick<EmailTemplate, 'name'|'subject'|'body'>>} updates
     */
    update(id, updates) {
        const store = loadStore();
        const t = store.templates.find((x) => x.id === id);
        if (!t) return null;
        if (updates.name != null) t.name = updates.name;
        if (updates.subject != null) t.subject = updates.subject;
        if (updates.body != null) t.body = updates.body;
        saveStore(store);
        return t;
    },

    /**
     * @param {string} id
     */
    remove(id) {
        const store = loadStore();
        const t = store.templates.find((x) => x.id === id);
        if (!t || t.builtin) return false;
        store.templates = store.templates.filter((x) => x.id !== id);
        if (store.defaultId === id) store.defaultId = store.templates[0]?.id || 'default';
        saveStore(store);
        return true;
    },

    setDefault(id) {
        const store = loadStore();
        if (!store.templates.some((x) => x.id === id)) return false;
        store.defaultId = id;
        saveStore(store);
        return true;
    },

    /**
     * Replace placeholders in subject/body with context values.
     * @param {{ subject: string; body: string }} template
     * @param {Object} ctx - { filename, date, signatureSummary, signerNames, pageCount, documentHash, attachmentNote }
     * @returns {{ subject: string; body: string }}
     */
    fill(template, ctx) {
        const attachmentNote =
            ctx.attachmentNote != null
                ? ctx.attachmentNote
                : `IMPORTANT: You must manually attach the PDF file (${ctx.filename || 'file'}) that was just downloaded to this email before sending. Do not attach a different version of the file.`;
        const map = {
            '{{filename}}': ctx.filename ?? '',
            '{{date}}': ctx.date ?? new Date().toLocaleString(),
            '{{signatureSummary}}': ctx.signatureSummary ?? 'No signatures.',
            '{{signerNames}}': ctx.signerNames ?? '—',
            '{{pageCount}}': String(ctx.pageCount ?? 0),
            '{{documentHash}}': ctx.documentHash ? `Document hash (SHA-256): ${ctx.documentHash}` : 'Document hash: N/A',
            '{{attachmentNote}}': attachmentNote
        };
        let subject = template.subject || '';
        let body = template.body || '';
        for (const [k, v] of Object.entries(map)) {
            subject = subject.split(k).join(v);
            body = body.split(k).join(v);
        }
        return { subject, body };
    },

    /**
     * Export all templates as JSON for file download / sync.
     * @returns {string}
     */
    exportJson() {
        const store = loadStore();
        return JSON.stringify({ ...store, exportedAt: new Date().toISOString() }, null, 2);
    },

    /**
     * Import templates from JSON (file content).
     * Merges with existing; duplicate ids overwrite. Set replace = true to replace all (keeps built-in default).
     * @param {string} json
     * @param {{ replace?: boolean }} options
     * @returns {{ imported: number; errors: string[] }}
     */
    importJson(json, options = {}) {
        const errors = [];
        let imported = 0;
        try {
            const data = JSON.parse(json);
            const incoming = Array.isArray(data.templates) ? data.templates : [];
            const store = options.replace
                ? { version: 1, defaultId: 'default', templates: [{ ...DEFAULT_TEMPLATE }] }
                : loadStore();

            for (const t of incoming) {
                if (!t || typeof t.id !== 'string' || !t.name) {
                    errors.push('Invalid template entry skipped.');
                    continue;
                }
                const isBuiltin = t.id === 'default' || !!t.builtin;
                const next = {
                    id: isBuiltin ? 'default' : t.id,
                    name: t.name,
                    subject: t.subject || '',
                    body: t.body || '',
                    isDefault: !!t.isDefault,
                    builtin: isBuiltin
                };
                const idx = store.templates.findIndex((x) => x.id === next.id);
                if (idx >= 0) store.templates[idx] = next;
                else store.templates.push(next);
                imported++;
            }

            if (data.defaultId && store.templates.some((x) => x.id === data.defaultId)) {
                store.defaultId = data.defaultId;
            }
            saveStore(store);
        } catch (e) {
            errors.push(e instanceof Error ? e.message : 'Invalid JSON');
        }
        return { imported, errors };
    }
};
