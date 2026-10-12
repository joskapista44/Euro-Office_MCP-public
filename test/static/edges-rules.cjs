'use strict';

const path = require('node:path');

/**
 * The M0 3.1 forbidden edges as pure predicates over a module list.
 *
 * They are separated from the file walk on purpose: today `src/adapters/` is empty (M1 migrates no
 * capability), so the adapter rules would pass on an empty subject set and prove nothing. Keeping
 * the rules pure lets the same code run against a synthetic fixture that MUST go red, which is
 * what makes the green on the real tree mean something.
 */

const SEP = path.sep;

function violations(modules) {
  const found = [];
  const add = (file, edge, rule) => found.push({ file, edge, rule });

  for (const mod of modules) {
    const isCore = mod.file.startsWith(`core${SEP}`) || mod.file.startsWith('core/');
    const isAdapter = mod.file.startsWith(`adapters${SEP}`) || mod.file.startsWith('adapters/');
    const own = isAdapter ? mod.file.split(/[\\/]/)[1] : null;

    for (const edge of mod.requires) {
      const e = edge.replace(/\\/g, '/');
      const f = mod.file.replace(/\\/g, '/');

      if (isCore && e.startsWith('adapters/')) add(f, e, 'core-imports-adapter');
      if (isAdapter && e.startsWith('core/session')) add(f, e, 'adapter-imports-session');
      if (isAdapter && e.startsWith('core/persistence')) add(f, e, 'adapter-imports-persistence');
      if (isAdapter && e.startsWith('adapters/') && e.split('/')[1] !== own) {
        add(f, e, 'adapter-imports-adapter');
      }
      if (f.startsWith('core/persistence/') && e.startsWith('core/receipt')) {
        add(f, e, 'persistence-imports-receipt');
      }
      if (f.startsWith('core/receipt/') && e.startsWith('core/persistence')) {
        add(f, e, 'receipt-imports-persistence');
      }

      // --- M2 rules (owner requirement 22) --------------------------------------------------
      // The observation layer is format-agnostic: it orchestrates an adapter through the
      // contract, and must never reach for one by name.
      if (f.startsWith('core/observation/') && e.startsWith('adapters/')) {
        add(f, e, 'observation-imports-adapter');
      }
      // Enrichment READS a package. It must not be able to dispatch, mutate, or own a session:
      // "package-only" must never become a write path.
      if (f.startsWith('core/observation/enrichment') && e.startsWith('core/session')) {
        add(f, e, 'enrichment-imports-session');
      }
      if (f.startsWith('core/observation/enrichment') && e.startsWith('core/execution')) {
        add(f, e, 'enrichment-imports-execution');
      }
      // A verifier reports what it saw. Persisting is not its concern, and a verifier that could
      // reach the barrier could hide a failure behind a save.
      if (f.startsWith('core/verification/') && e.startsWith('core/persistence')) {
        add(f, e, 'verifier-imports-persistence');
      }
      // The authority lattice and the mutation epoch are LEAF modules: everything may depend on
      // them, and they may depend on nothing, or the graph acquires a cycle.
      if ((f === 'core/authority.cjs' || f === 'core/mutation-epoch.cjs') && e.startsWith('core/')) {
        add(f, e, 'leaf-module-has-an-edge');
      }
    }
  }
  return found;
}

module.exports = { violations };
