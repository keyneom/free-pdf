/**
 * Signing metadata - Parse and build our app's signing-flow payload stored in PDF Keywords.
 * Format: Keywords = "free-pdf-v1 " + base64(JSON.stringify(payload))
 * Payload: { v: 1, signers: [...], expectedSigners?: [...], lockedSignatureFields?: string[], documentStage?: 'draft'|'sent'|'signed', hashChain?: { h, t, p }, ... }
 * - lockedSignatureFields: field labels that are signed and must not be modified by another participant.
 * - documentStage: draft (editable), sent (sent for signing), signed (has signatures; treat as received).
 * - hashChain: latest link { h: documentHash, t: timestamp, p: previousHash } for proving when changes were made.
 */

const PREFIX = 'free-pdf-v1 ';

function parseEmailList(arr) {
    return Array.isArray(arr)
        ? arr.filter((e) => typeof e === 'string' && e.trim()).map((e) => String(e).trim())
        : undefined;
}

/**
 * Parse Keywords string. Returns null if not our metadata.
 * @param {string} [keywords]
 * @returns {{ signers: Array<...>, expectedSigners?: Array<...>, lockedSignatureFields?: string[], documentStage?: string, hashChain?: { hash, timestamp, previousHash }, ... } | null}
 */
export function parseSigningMetadata(keywords) {
    if (typeof keywords !== 'string' || !keywords.startsWith(PREFIX)) return null;
    try {
        const raw = keywords.slice(PREFIX.length).trim();
        if (!raw) return { signers: [], expectedSigners: [], emailTemplate: undefined, originalSenderEmail: undefined, completionToEmails: undefined, completionCcEmails: undefined, completionBccEmails: undefined, lockedSignatureFields: [], lockedFormFields: [], documentStage: undefined, hashChain: undefined };
        const json = decodeURIComponent(escape(atob(raw)));
        const data = JSON.parse(json);
        if (!data || data.v !== 1) return null;
        const signers = (data.signers || []).map((s) => ({
            name: s.n || '',
            timestamp: s.t || ''
        }));
        const expectedSigners = (data.expectedSigners || []).map((e, i) => ({
            name: e.n || '',
            email: e.e,
            order: typeof e.o === 'number' ? e.o : i + 1
        }));
        const emailTemplate =
            data.emailTemplate && typeof data.emailTemplate.subject === 'string' && typeof data.emailTemplate.body === 'string'
                ? { subject: data.emailTemplate.subject, body: data.emailTemplate.body }
                : undefined;
        const originalSenderEmail = typeof data.originalSenderEmail === 'string' && data.originalSenderEmail.trim() ? data.originalSenderEmail.trim() : undefined;
        const completionToEmails = parseEmailList(data.completionToEmails);
        const completionCcEmails = parseEmailList(data.completionCcEmails);
        const completionBccEmails = parseEmailList(data.completionBccEmails);
        const lockedSignatureFields = Array.isArray(data.lockedSignatureFields) ? data.lockedSignatureFields.filter((l) => typeof l === 'string').map((l) => String(l).trim()) : [];
        const lockedFormFields = Array.isArray(data.lockedFormFields) ? data.lockedFormFields.filter((l) => typeof l === 'string').map((l) => String(l).trim()) : [];
        const documentStage = typeof data.documentStage === 'string' && ['draft', 'sent', 'signed'].includes(data.documentStage) ? data.documentStage : undefined;
        const hashChain = data.hashChain && typeof data.hashChain === 'object' && typeof data.hashChain.h === 'string'
            ? { hash: data.hashChain.h, timestamp: data.hashChain.t || '', previousHash: typeof data.hashChain.p === 'string' ? data.hashChain.p : undefined }
            : undefined;
        return { signers, expectedSigners, emailTemplate, originalSenderEmail, completionToEmails, completionCcEmails, completionBccEmails, lockedSignatureFields, lockedFormFields, documentStage, hashChain };
    } catch {
        return null;
    }
}

/**
 * Build Keywords string for embedding in PDF.
 * @param {{ signers: Array<...>, expectedSigners?: Array<...>, lockedSignatureFields?: string[], documentStage?: 'draft'|'sent'|'signed', hashChain?: { hash, timestamp, previousHash }, ... }} payload
 * @returns {string}
 */
export function buildSigningKeywords(payload) {
    if (!payload || !Array.isArray(payload.signers)) return '';
    const data = {
        v: 1,
        signers: payload.signers.map((s) => ({ n: s.name || '', t: s.timestamp || '' })),
        expectedSigners: (payload.expectedSigners || []).map((e, i) => ({
            n: e.name || '',
            e: e.email,
            o: typeof e.order === 'number' ? e.order : i + 1
        }))
    };
    if (payload.emailTemplate && typeof payload.emailTemplate.subject === 'string' && typeof payload.emailTemplate.body === 'string') {
        data.emailTemplate = { subject: payload.emailTemplate.subject, body: payload.emailTemplate.body };
    }
    if (typeof payload.originalSenderEmail === 'string' && payload.originalSenderEmail.trim()) {
        data.originalSenderEmail = payload.originalSenderEmail.trim();
    }
    if (Array.isArray(payload.completionToEmails) && payload.completionToEmails.length > 0) {
        data.completionToEmails = payload.completionToEmails.filter((e) => typeof e === 'string' && e.trim()).map((e) => String(e).trim());
    }
    if (Array.isArray(payload.completionCcEmails) && payload.completionCcEmails.length > 0) {
        data.completionCcEmails = payload.completionCcEmails.filter((e) => typeof e === 'string' && e.trim()).map((e) => String(e).trim());
    }
    if (Array.isArray(payload.completionBccEmails) && payload.completionBccEmails.length > 0) {
        data.completionBccEmails = payload.completionBccEmails.filter((e) => typeof e === 'string' && e.trim()).map((e) => String(e).trim());
    }
    if (Array.isArray(payload.lockedSignatureFields) && payload.lockedSignatureFields.length > 0) {
        data.lockedSignatureFields = payload.lockedSignatureFields.filter((l) => typeof l === 'string' && l.trim()).map((l) => String(l).trim());
    }
    if (Array.isArray(payload.lockedFormFields) && payload.lockedFormFields.length > 0) {
        data.lockedFormFields = payload.lockedFormFields.filter((l) => typeof l === 'string' && l.trim()).map((l) => String(l).trim());
    }
    if (typeof payload.documentStage === 'string' && ['draft', 'sent', 'signed'].includes(payload.documentStage)) {
        data.documentStage = payload.documentStage;
    }
    if (payload.hashChain && typeof payload.hashChain.hash === 'string') {
        data.hashChain = {
            h: payload.hashChain.hash,
            t: typeof payload.hashChain.timestamp === 'string' ? payload.hashChain.timestamp : new Date().toISOString(),
            p: typeof payload.hashChain.previousHash === 'string' ? payload.hashChain.previousHash : undefined
        };
    }
    return PREFIX + btoa(unescape(encodeURIComponent(JSON.stringify(data))));
}

/**
 * Check if a PDF has our metadata (Producer or Keywords).
 * @param {{ info?: { Keywords?: string, Producer?: string } }} metadata - Result of getMetadata()
 * @returns {boolean}
 */
export function hasOurSigningMetadata(metadata) {
    const info = metadata?.info || {};
    if (typeof info.Keywords === 'string' && info.Keywords.startsWith(PREFIX)) return true;
    if (typeof info.Producer === 'string' && info.Producer.includes('PDF Editor')) return true;
    return false;
}

/**
 * Compute SHA-256 hash of a string for the document hash chain.
 * @param {string} data - Canonical string (e.g. JSON of viewPages + annotation digests)
 * @returns {Promise<string>} Hex-encoded hash
 */
export async function computeDocumentHash(data) {
    if (typeof crypto !== 'object' || !crypto.subtle) {
        return '';
    }
    try {
        const encoder = new TextEncoder();
        const bytes = encoder.encode(typeof data === 'string' ? data : JSON.stringify(data));
        const hashBuffer = await crypto.subtle.digest('SHA-256', bytes);
        const hashArray = Array.from(new Uint8Array(hashBuffer));
        return hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
    } catch {
        return '';
    }
}
