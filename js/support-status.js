/**
 * Support status: first document free, then time-based OR use-based expiration.
 * User chooses one: either time period OR document uses, not both.
 * Amount scales: $1 = 7 days or 5 uses, scales linearly (e.g. $5 = 35 days or 25 uses)
 */

const KEY_FIRST_USED = 'freePdfFirstDocumentUsed';
const KEY_SUPPORT = 'freePdfSupportStatus';

/** @typedef {'time'|'uses'} SupportMode */

/**
 * @returns {boolean} true if the user has used download/send at least once
 */
export function isFirstDocumentUsed() {
    try {
        return localStorage.getItem(KEY_FIRST_USED) === 'true';
    } catch {
        return false;
    }
}

/**
 * Mark that the user has used download/send for the first time.
 */
export function markFirstDocumentUsed() {
    try {
        localStorage.setItem(KEY_FIRST_USED, 'true');
    } catch (_) {}
}

/**
 * Format days as human-readable string (e.g. "1 week", "2 weeks", "1 month").
 * @param {number} days
 * @returns {string}
 */
export function formatDuration(days) {
    const d = Math.max(1, Math.floor(days));
    if (d >= 30) {
        const months = Math.floor(d / 30);
        return months === 1 ? '1 month' : `${months} months`;
    }
    if (d >= 7) {
        const weeks = Math.floor(d / 7);
        return weeks === 1 ? '1 week' : `${weeks} weeks`;
    }
    return d === 1 ? '1 day' : `${d} days`;
}

/**
 * @param {number} amountUsd - any positive amount
 * @returns {number} duration in days (~7 per dollar)
 */
export function getDurationDaysForAmount(amountUsd) {
    const amt = Math.max(0, Number(amountUsd) || 0);
    return Math.max(1, Math.floor(amt * 7));
}

/**
 * @param {number} amountUsd - any positive amount
 * @returns {number} number of download/send uses (~5 per dollar)
 */
export function getUsesForAmount(amountUsd) {
    const amt = Math.max(0, Number(amountUsd) || 0);
    return Math.max(1, Math.floor(amt * 5));
}

/**
 * Record a payment (call when user returns from payment flow).
 * @param {number} amountUsd - any positive amount
 * @param {SupportMode} mode - 'time' (duration only) or 'uses' (download/send count only)
 */
export function recordSupportDonation(amountUsd, mode = 'time') {
    try {
        const amt = Math.max(0.5, Number(amountUsd) || 1);
        const data = {
            lastDonationAt: Date.now(),
            amountUsd: amt,
            mode: mode === 'uses' ? 'uses' : 'time',
            usesRemaining: mode === 'uses' ? getUsesForAmount(amt) : 0
        };
        localStorage.setItem(KEY_SUPPORT, JSON.stringify(data));
    } catch (_) {}
}

/**
 * Consume one use when user completes a download or send.
 * No-op if no support record, mode is 'time', or uses already 0.
 */
export function consumeSupportUse() {
    try {
        const raw = localStorage.getItem(KEY_SUPPORT);
        if (!raw) return;

        const data = JSON.parse(raw);
        if (data?.mode !== 'uses') return;

        let uses = data?.usesRemaining;
        if (uses == null || uses <= 0) return;

        data.usesRemaining = uses - 1;
        localStorage.setItem(KEY_SUPPORT, JSON.stringify(data));
    } catch (_) {}
}

/**
 * @returns {{ valid: boolean, lastDonationAt?: number, expiredAt?: number, usesRemaining?: number, mode?: SupportMode }}
 */
export function getSupportStatus() {
    try {
        const raw = localStorage.getItem(KEY_SUPPORT);
        if (!raw) return { valid: false };

        const data = JSON.parse(raw);
        const lastDonationAt = data?.lastDonationAt;
        const amountUsd = data?.amountUsd ?? 1;
        const mode = data?.mode === 'uses' ? 'uses' : 'time';
        const usesRemaining = Math.max(0, data?.usesRemaining ?? 0);

        if (mode === 'time') {
            if (!lastDonationAt || typeof lastDonationAt !== 'number') return { valid: false };
            const days = getDurationDaysForAmount(amountUsd);
            const expiryMs = lastDonationAt + days * 24 * 60 * 60 * 1000;
            const valid = Date.now() < expiryMs;
            return { valid, lastDonationAt, expiredAt: expiryMs, usesRemaining: 0, mode };
        } else {
            const valid = usesRemaining > 0;
            return {
                valid,
                lastDonationAt: lastDonationAt ?? 0,
                expiredAt: 0,
                usesRemaining,
                mode
            };
        }
    } catch {
        return { valid: false };
    }
}

/**
 * @returns {boolean} true if support is still valid (within time or uses remaining)
 */
export function isSupportValid() {
    return getSupportStatus().valid;
}
