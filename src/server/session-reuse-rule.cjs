'use strict';

/**
 * GH #100 U3 -- the REUSE RULE: may a kept (warm) editor session serve the next call? A pure function of what the lifecycle measured at call start;
 * it opens nothing, fetches nothing and decides nothing about limits (idle / age / calls are the registry's, session-registry.cjs).
 * docs/GH100-U3-FRESHNESS-PLAN.md has the measurement behind every check.
 *
 *   MEASURED (LIVE probe, prod, broker transport): the DS document key of the editor config is constant for the life of ONE document session --
 *   through the session's own save and a foreign editor's save -- and it CHANGED once that session had closed (a fresh open got the new key).
 *   So it cannot see an in-session change, and it does not need to: a foreign EDITOR's write reached the warm session through the DS co-editing
 *   socket (the warm session read the foreign value). What it CAN see is that the document session this editor is attached to is no longer
 *   the current one (a rolled key at call start) -- the one signal that says "do not trust this model".
 *   NOT MEASURED: a non-editor WebDAV write while the session is warm (the broker transport has no PUT). It is expected to be refused by the lock
 *   (derived, not measured); reuse therefore ships OFF, and enabling it needs that case measured.
 *
 * Every unknown is a "no": a missing, null or non-boolean input never reads as healthy.
 *
 * GH #100 U4c -- operation axis (docs/GH100-U4C-READ-REUSE-DESIGN.md section 4): a held session was
 * opened for READ or WRITE (recorded at hold time); a WRITE request may only take a session that was
 * itself opened for WRITE (an edit-capable session), because a READ-mode config may have been served
 * view-only (broker-config-source.cjs). A READ request accepts either -- an edit-capable session can
 * always serve what a view-only one could. Checked FIRST (cheapest), before any health/freshness I/O.
 */

function decideReuse({ health, openedKey, freshKey, lastRelease, requestedOperation, openedOperation } = {}) {
  const no = reason => Object.freeze({ reuse: false, reason });
  if (requestedOperation === 'WRITE' && openedOperation !== 'WRITE') return no('held-session-not-write-capable');
  if (!health || health.alive !== true) return no('editor-not-alive');
  if (health.canSave !== false) return no(health.canSave === true ? 'unsaved-changes-in-editor' : 'cansave-unreadable');
  const ds = health.ds;
  if (!ds || ds.open !== true) return no('ds-socket-not-open');
  if (ds.busy !== false) return no('save-in-flight');
  if (health.anomalies !== 0) return no('observer-anomaly');
  const rel = lastRelease;
  const clean = rel && rel.ok === true && (rel.confirmation === null || rel.confirmation === undefined || (rel.confirmation && rel.confirmation.state === 'confirmed'));
  if (!clean) return no('last-call-end-not-clean');
  if (typeof openedKey !== 'string' || openedKey === '') return no('no-recorded-document-key');
  if (typeof freshKey !== 'string' || freshKey === '') return no('fresh-config-key-unavailable');
  if (openedKey !== freshKey) return no('document-session-changed');
  return Object.freeze({ reuse: true, reason: 'all-checks-passed' });
}

module.exports = { decideReuse };
