/**
 * Secure Storage - Password-protected vaults for templates, signatures, etc.
 * Multiple named vaults; each encrypted with its own password. Web Crypto (PBKDF2 + AES-GCM).
 */

const REGISTRY_KEY = 'free-pdf-vault-registry';
const VAULT_PREFIX = 'free-pdf-vault-';
const EMAIL_TEMPLATES_KEY = 'free-pdf-email-templates';

// Legacy single-vault keys (for migration)
const LEGACY_META_KEY = 'free-pdf-vault-meta';
const LEGACY_VAULT_KEY = 'free-pdf-secure-vault';

const PBKDF2_ITERATIONS = 230_000;
const SALT_LEN = 16;
const IV_LEN = 12;
const KEY_LEN = 256;

let _key = null;
let _vault = null;
let _activeVaultId = null;

function b64enc(u8) {
    return btoa(String.fromCharCode.apply(null, u8));
}

function b64dec(s) {
    const bin = atob(s);
    const u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8;
}

function vaultId() {
    return 'vault-' + Math.random().toString(36).slice(2, 10);
}

async function deriveKey(password, salt) {
    const enc = new TextEncoder();
    const keyMaterial = await crypto.subtle.importKey(
        'raw',
        enc.encode(password),
        'PBKDF2',
        false,
        ['deriveBits', 'deriveKey']
    );
    return crypto.subtle.deriveKey(
        {
            name: 'PBKDF2',
            salt,
            iterations: PBKDF2_ITERATIONS,
            hash: 'SHA-256'
        },
        keyMaterial,
        { name: 'AES-GCM', length: KEY_LEN },
        false,
        ['encrypt', 'decrypt']
    );
}

async function encrypt(plaintext, key) {
    const iv = crypto.getRandomValues(new Uint8Array(IV_LEN));
    const enc = new TextEncoder();
    const pt = enc.encode(plaintext);
    const ct = await crypto.subtle.encrypt(
        { name: 'AES-GCM', iv, tagLength: 128 },
        key,
        pt
    );
    const combined = new Uint8Array(IV_LEN + ct.byteLength);
    combined.set(iv);
    combined.set(new Uint8Array(ct), IV_LEN);
    return b64enc(combined);
}

async function decrypt(b64, key) {
    const combined = b64dec(b64);
    const iv = combined.slice(0, IV_LEN);
    const ct = combined.slice(IV_LEN);
    const buf = await crypto.subtle.decrypt(
        { name: 'AES-GCM', iv, tagLength: 128 },
        key,
        ct
    );
    return new TextDecoder().decode(buf);
}

function getRegistry() {
    try {
        const raw = localStorage.getItem(REGISTRY_KEY);
        if (!raw) return [];
        const arr = JSON.parse(raw);
        return Array.isArray(arr) ? arr : [];
    } catch (_) {
        return [];
    }
}

function saveRegistry(registry) {
    localStorage.setItem(REGISTRY_KEY, JSON.stringify(registry));
}

function loadVaultPayload(id) {
    return localStorage.getItem(VAULT_PREFIX + id);
}

function saveVaultPayload(id, b64) {
    localStorage.setItem(VAULT_PREFIX + id, b64);
}

function defaultTemplates() {
    return {
        version: 1,
        defaultId: 'default',
        templates: [{
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
        }]
    };
}

function migrateTemplatesFromLegacy() {
    const raw = localStorage.getItem(EMAIL_TEMPLATES_KEY);
    if (!raw) return null;
    try {
        const data = JSON.parse(raw);
        if (data && typeof data.version === 'number' && Array.isArray(data.templates)) {
            return data;
        }
    } catch (_) {}
    return null;
}

function removeLegacyTemplates() {
    localStorage.removeItem(EMAIL_TEMPLATES_KEY);
}

function uid() {
    return 'sig-' + Math.random().toString(36).slice(2, 12);
}

/**
 * Migrate from legacy single-vault format to multi-vault registry.
 * Preserves the same encrypted blob and salt so the user's password still works.
 */
function migrateFromLegacyIfNeeded() {
    const registry = getRegistry();
    if (registry.length > 0) return;

    const oldMeta = (() => {
        try {
            const raw = localStorage.getItem(LEGACY_META_KEY);
            if (!raw) return null;
            const m = JSON.parse(raw);
            return m?.salt ? m : null;
        } catch (_) {
            return null;
        }
    })();
    const oldBlob = localStorage.getItem(LEGACY_VAULT_KEY);
    if (!oldMeta || !oldBlob) return;

    const id = vaultId();
    const entry = {
        id,
        name: 'Default',
        salt: oldMeta.salt,
        createdAt: new Date().toISOString()
    };
    saveVaultPayload(id, oldBlob);
    saveRegistry([entry]);
    localStorage.removeItem(LEGACY_META_KEY);
    localStorage.removeItem(LEGACY_VAULT_KEY);
}

export const secureStorage = {
    hasVault() {
        return getRegistry().length > 0;
    },

    isUnlocked() {
        return _key != null && _vault != null && _activeVaultId != null;
    },

    getRegistry() {
        return [...getRegistry()];
    },

    getActiveVaultId() {
        return _activeVaultId ?? null;
    },

    getActiveVaultName() {
        if (!_activeVaultId) return '';
        const reg = getRegistry().find((r) => r.id === _activeVaultId);
        return reg?.name ?? '';
    },

    /**
     * Create a new named vault. Migrates legacy templates into the first vault if present.
     * @param {string} name - Display name for the vault
     * @param {string} password
     */
    async createVault(name, password) {
        const registry = getRegistry();
        const trimmedName = (name || '').trim() || 'Unnamed';
        const salt = crypto.getRandomValues(new Uint8Array(SALT_LEN));
        const key = await deriveKey(password, salt);

        const legacy = registry.length === 0 ? migrateTemplatesFromLegacy() : null;
        const templates = legacy || defaultTemplates();
        if (!templates.templates?.length) {
            templates.templates = defaultTemplates().templates;
            templates.defaultId = 'default';
        }
        const vault = { version: 1, templates, signatures: [] };
        const b64 = await encrypt(JSON.stringify(vault), key);

        const id = vaultId();
        const entry = { id, name: trimmedName, salt: b64enc(salt), createdAt: new Date().toISOString() };
        saveVaultPayload(id, b64);
        saveRegistry([...registry, entry]);
        if (legacy) removeLegacyTemplates();

        _activeVaultId = id;
        _key = key;
        _vault = vault;
    },

    /**
     * Unlock a vault by id with password.
     * @param {string} id - Vault id from registry
     * @param {string} password
     */
    async unlock(id, password) {
        const registry = getRegistry();
        const entry = registry.find((r) => r.id === id);
        if (!entry) throw new Error('Vault not found.');
        const b64 = loadVaultPayload(id);
        if (!b64) throw new Error('Vault data not found.');
        const salt = b64dec(entry.salt);
        const key = await deriveKey(password, salt);
        try {
            const raw = await decrypt(b64, key);
            _vault = JSON.parse(raw);
        } catch (e) {
            throw new Error('Wrong password or corrupted vault.');
        }
        _activeVaultId = id;
        _key = key;
    },

    lock() {
        _key = null;
        _vault = null;
        _activeVaultId = null;
    },

    /**
     * Verify password for a vault (derive key and decrypt); throws if wrong.
     * @param {string} id - Vault id
     * @param {string} password
     */
    async verifyPassword(id, password) {
        const registry = getRegistry();
        const entry = registry.find((r) => r.id === id);
        if (!entry) throw new Error('Vault not found.');
        const b64 = loadVaultPayload(id);
        if (!b64) throw new Error('Vault data not found.');
        const salt = b64dec(entry.salt);
        const key = await deriveKey(password, salt);
        try {
            await decrypt(b64, key);
        } catch (e) {
            throw new Error('Wrong password.');
        }
    },

    /**
     * Delete a vault. Requires password to prove ownership. Removes from registry and deletes payload.
     * @param {string} id - Vault id
     * @param {string} password
     */
    async deleteVault(id, password) {
        await this.verifyPassword(id, password);
        const registry = getRegistry().filter((r) => r.id !== id);
        saveRegistry(registry);
        localStorage.removeItem(VAULT_PREFIX + id);
        if (_activeVaultId === id) this.lock();
    },

    /**
     * Rename a vault. Requires password to prove ownership. Only updates display name (id unchanged; no sync impact).
     * @param {string} id - Vault id
     * @param {string} password
     * @param {string} newName
     */
    async renameVault(id, password, newName) {
        await this.verifyPassword(id, password);
        const trimmed = (newName || '').trim() || 'Unnamed';
        const registry = getRegistry().map((r) => (r.id === id ? { ...r, name: trimmed } : r));
        saveRegistry(registry);
    },

    /**
     * Export the currently unlocked vault for backup or transfer to another device.
     * @returns {{ version: number; name: string; salt: string; payload: string; exportedAt: string }}
     */
    exportVault() {
        this._assertUnlocked();
        const entry = getRegistry().find((r) => r.id === _activeVaultId);
        if (!entry) throw new Error('Vault not found.');
        const payload = loadVaultPayload(_activeVaultId);
        if (!payload) throw new Error('Vault data not found.');
        return {
            version: 1,
            name: entry.name,
            salt: entry.salt,
            payload,
            exportedAt: new Date().toISOString()
        };
    },

    /**
     * Import a vault from an exported file as a new vault. Same password works on both devices.
     * @param {{ name: string; salt: string; payload: string }} fileData
     * @param {string} password - Password for the imported vault (used to verify; vault stores same salt/payload)
     */
    async importVaultAsNew(fileData, password) {
        if (!fileData?.name || !fileData?.salt || !fileData?.payload) throw new Error('Invalid vault file.');
        const salt = b64dec(fileData.salt);
        const key = await deriveKey(password, salt);
        try {
            await decrypt(fileData.payload, key);
        } catch (e) {
            throw new Error('Wrong password for this vault file.');
        }
        const registry = getRegistry();
        const names = new Set(registry.map((r) => r.name));
        let name = (fileData.name || '').trim() || 'Imported';
        if (names.has(name)) {
            let n = 1;
            while (names.has(`${name} (${n})`)) n++;
            name = `${name} (${n})`;
        }
        const id = vaultId();
        const entry = { id, name, salt: fileData.salt, createdAt: new Date().toISOString() };
        saveVaultPayload(id, fileData.payload);
        saveRegistry([...registry, entry]);
        _activeVaultId = id;
        _key = await deriveKey(password, salt);
        const raw = await decrypt(fileData.payload, _key);
        _vault = JSON.parse(raw);
    },

    /**
     * Replace the currently unlocked vault's content with data from an imported file.
     * @param {{ salt: string; payload: string }} fileData
     * @param {string} filePassword - Password for the vault file
     */
    async replaceVaultWithImport(fileData, filePassword) {
        this._assertUnlocked();
        if (!fileData?.salt || !fileData?.payload) throw new Error('Invalid vault file.');
        const salt = b64dec(fileData.salt);
        const key = await deriveKey(filePassword, salt);
        let data;
        try {
            const raw = await decrypt(fileData.payload, key);
            data = JSON.parse(raw);
        } catch (e) {
            throw new Error('Wrong password for this vault file.');
        }
        _vault = data;
        const b64 = await encrypt(JSON.stringify(_vault), _key);
        saveVaultPayload(_activeVaultId, b64);
    },

    _assertUnlocked() {
        if (!this.isUnlocked()) throw new Error('Vault is locked. Unlock first.');
    },

    getTemplatesStore() {
        this._assertUnlocked();
        return _vault.templates;
    },

    async saveTemplatesStore(store) {
        this._assertUnlocked();
        _vault.templates = store;
        const b64 = await encrypt(JSON.stringify(_vault), _key);
        saveVaultPayload(_activeVaultId, b64);
    },

    getSignatures() {
        this._assertUnlocked();
        return _vault.signatures || [];
    },

    async addSignature(entry) {
        this._assertUnlocked();
        const list = _vault.signatures || [];
        const next = { id: uid(), name: entry.name || 'Untitled', dataUrl: entry.dataUrl, type: entry.type || 'draw', createdAt: new Date().toISOString() };
        list.push(next);
        _vault.signatures = list;
        const b64 = await encrypt(JSON.stringify(_vault), _key);
        saveVaultPayload(_activeVaultId, b64);
        return next;
    },

    async removeSignature(id) {
        this._assertUnlocked();
        const list = (_vault.signatures || []).filter((s) => s.id !== id);
        _vault.signatures = list;
        const b64 = await encrypt(JSON.stringify(_vault), _key);
        saveVaultPayload(_activeVaultId, b64);
    },

    /** Call once on app load to migrate from legacy single-vault format. */
    migrateFromLegacyIfNeeded() {
        migrateFromLegacyIfNeeded();
    }
};
