# Rollback

## Product

Point `current` back at the previous release directory and restart the MCP clients' server processes:

```sh
ln -sfn releases/<previous ARTIFACT_VERSION> /opt/euro-office-mcp/current
/opt/euro-office-mcp/current/bin/euro-office-mcp.cjs --identity     # must show the previous identity
```

Then the same smoke as after an update. No data migration is involved: apart from its own log file the server keeps no state.

## Document Server blocks

Restore the bundle copy kept before the install ([INSTALL.md](INSTALL.md) §1) atomically, keeping owner, group and
mode, and verify the served bundle's hash. Only roll the Document Server back to a block set the currently served
release accepts -- otherwise roll the product back first. A release whose required capability is not served refuses
that operation with a named reason; it does not fall back to anything else.

## Prove it before you need it

Rehearse rollback and forward-restore once per release on a non-production instance: identity before and after, and
the smoke in both directions.
