/**
 * Support status: first document free, then time-based OR use-based expiration.
 * User chooses one: either time period OR document uses, not both.
 * Tier: $1=7d or 5 uses, $5=30d or 25 uses, $20=180d or 100 uses
 */

const KEY_FIRST_USED = 'freePdfFirstDocumentUsed';
const KEY_SUPPORT = 'freePdfSupportStatus';

const TIER_DAYS = {
    1: 7,
    5: 30,
    20: 180
};

const TIER_USES = {
    1: 5,
    5: 25,
    20: 100
};

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
 * @param {number} amountUsd - 1, 5, or 20
 * @returns {number} duration in days
 */
export function getDurationDaysForAmount(amountUsd) {
    return TIER_DAYS[amountUsd] ?? 7;
}

/**
 * @param {number} amountUsd - 1, 5, or 20
 * @returns {number} number of download/send uses
 */
export function getUsesForAmount(amountUsd) {
    return TIER_USES[amountUsd] ?? 5;
}

/**
 * Record a donation (call when user returns from donation flow).
 * @param {number} amountUsd - 1, 5, or 20
 * @param {SupportMode} mode - 'time' (duration only) or 'uses' (download/send count only)
 */
export function recordSupportDonation(amountUsd, mode = 'time') {
    try {
        const amt = Number(amountUsd) || 1;
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
