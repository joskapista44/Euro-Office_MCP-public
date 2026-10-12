'use strict';

const { singletonGroups } = require('./plan.cjs');

/**
 * Sequential strategy — M3 point 8, the correctness baseline.
 *
 * Every operation is its own group, so each one completes classify -> resolve -> dispatch ->
 * fresh read -> verify before the next mutation dispatch starts. Execution itself lives in the
 * common engine (core/execution/executor.cjs); a strategy only decides the grouping.
 */
module.exports = Object.freeze({
  name: 'sequential',
  groups: operations => singletonGroups(operations),
  transaction: false,
});
