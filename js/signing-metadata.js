/**
 * Signing metadata - Parse and build our app's signing-flow payload stored in PDF Keywords.
 * Format: Keywords = "free-pdf-v1 " + base64(JSON.stringify(payload))
 * Payload: { v: 1, signers: [...], expectedSigners?: [{ name/fieldLabel, email, order }], ... }
 * expectedSigners[].name = signature field label (which slot this signer fills), not the person's legal name.
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
 * @returns {{ signers: Array<...>, expectedSigners?: Array<...>, emailTemplate?: { subject, body }, originalSenderEmail?: string, completionToEmails?: string[], completionCcEmails?: string[], completionBccEmails?: string[] } | null}
 */
export function parseSigningMetadata(keywords) {
    if (typeof keywords !== 'string' || !keywords.startsWith(PREFIX)) return null;
    try {
        const raw = keywords.slice(PREFIX.length).trim();
        if (!raw) return { signers: [], expectedSigners: [], emailTemplate: undefined, originalSenderEmail: undefined, completionToEmails: undefined, completionCcEmails: undefined, completionBccEmails: undefined };
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
        return { signers, expectedSigners, emailTemplate, originalSenderEmail, completionToEmails, completionCcEmails, completionBccEmails };
    } catch {
        return null;
    }
}

/**
 * Build Keywords string for embedding in PDF.
 * @param {{ signers: Array<{ name: string, timestamp: string }>, expectedSigners?: Array<{ name: string, email?: string, order?: number }>, emailTemplate?: { subject: string, body: string }, originalSenderEmail?: string, completionToEmails?: string[] }} payload
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
    if (typeof info.Producer === 'string' && info.Producer.includes('Free PDF Editor')) return true;
    return false;
}
