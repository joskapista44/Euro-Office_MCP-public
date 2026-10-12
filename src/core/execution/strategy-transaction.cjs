'use strict';

const { singletonGroups } = require('./plan.cjs');

/**
 * Transaction strategy — M3 point 9, M0 section 10.
 *
 *   TRANSACTION_ATOMICITY = NON_ATOMIC
 *
 * Ordered, stop on the first failure, no further dispatch, NO rollback: verified writes before
 * the failure stay in the document and route through the M1 barrier. It is not all-or-nothing and
 * must never be described as such. The difference from `sequential` is the declared contract,
 * reported in ExecutionResult.transaction, not a second execution path.
 */
module.exports = Object.freeze({
  name: 'transaction',
  groups: operations => singletonGroups(operations),
  transaction: Object.freeze({ atomicity: 'NON_ATOMIC', stopOnFirstFailure: true, rollback: false }),
});
