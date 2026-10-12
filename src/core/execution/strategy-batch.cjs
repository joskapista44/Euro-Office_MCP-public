'use strict';

const { groupIntoBatches } = require('./plan.cjs');

/**
 * Batch strategy — M3 points 10-13, M0 section 10.1.
 *
 * Fewer interactions, never weaker correctness. Independent operations share one targeting
 * observation, one dispatch round and one post-dispatch read; a dependency ALWAYS ends the group
 * (plan.groupIntoBatches), so a dependent operation is targeted from the state its dependency
 * produced. This module decides the grouping and nothing else: it cannot dispatch, read or verify,
 * which is what keeps a "fast path" from skipping a barrier -- there is no path here to skip with.
 */
module.exports = Object.freeze({
  name: 'batch',
  groups: (operations, options = {}) => groupIntoBatches(operations, options),
  transaction: false,
});
