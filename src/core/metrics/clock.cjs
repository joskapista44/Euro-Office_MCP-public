'use strict';

/**
 * Named clock boundaries — M0 section 12.1.
 *
 * The two sources measured different spans under similar names, so the boundaries are named here
 * and nowhere else. `taskElapsedMs` keeps the XLSX meaning (open excluded) and `totalElapsedMs`
 * keeps the DOCX meaning (open included), so both historical series stay interpretable.
 *
 *   sessionOpenMs   openSession entry  -> handle returned
 *   readyMs         handle returned    -> READY reached
 *   taskElapsedMs   READY reached      -> executeTask returned
 *   totalElapsedMs  openSession entry  -> closeSession returned
 */

const BOUNDARIES = Object.freeze({
  SESSION_OPEN: 'sessionOpenMs',
  READY: 'readyMs',
  TASK_ELAPSED: 'taskElapsedMs',
  TOTAL_ELAPSED: 'totalElapsedMs',
});

/** The default source of time. Injectable so tests are deterministic rather than timing-dependent. */
const defaultNow = () => Date.now();

class Clock {
  constructor({ now = defaultNow } = {}) {
    this._now = now;
    this._marks = new Map();
    this._spans = new Map();
  }

  /** Records a named instant. Re-marking the same name is a caller error and throws. */
  mark(name) {
    if (this._marks.has(name)) {
      throw new Error(`clock mark already taken: ${name}`);
    }
    this._marks.set(name, this._now());
    return this;
  }

  has(name) {
    return this._marks.has(name);
  }

  /** The current instant from the injected source. Callers measuring their own span use this. */
  now() {
    return this._now();
  }

  /**
   * Closes a boundary between two marks. A span whose endpoints were never marked stays absent
   * rather than defaulting to 0 — an unmeasured duration must not read as an instant one.
   */
  span(boundary, fromMark, toMark) {
    if (!this._marks.has(fromMark) || !this._marks.has(toMark)) return this;
    this._spans.set(boundary, this._marks.get(toMark) - this._marks.get(fromMark));
    return this;
  }

  /** Spans that were never closed are null, and the metric schema reports them as null. */
  spans() {
    const out = {};
    for (const boundary of Object.values(BOUNDARIES)) {
      out[boundary] = this._spans.has(boundary) ? this._spans.get(boundary) : null;
    }
    return out;
  }
}

module.exports = { Clock, BOUNDARIES, defaultNow };
