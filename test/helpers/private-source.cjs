'use strict';

/**
 * The public tree is derived from the canonical private repository and does NOT carry the LIVE harnesses
 * (`test/live/`, they need the private broker and the internal Document Server) or the internal docs. A test CASE
 * whose subject is one of those private files is skipped there with a named reason; in the canonical repository the
 * file exists and the case runs unchanged. Only the presence of the named path decides -- never an environment flag.
 */
const fs = require('node:fs');
const path = require('node:path');

const ROOT = path.join(__dirname, '..', '..');

/** false when every named repository path exists (the case runs), else the named skip reason. */
const needsPrivate = (...rels) => {
  const missing = rels.filter(rel => !fs.existsSync(path.join(ROOT, rel)));
  return missing.length ? `needs ${missing.join(', ')} -- private source, not part of the public tree` : false;
};

module.exports = { needsPrivate };
