# Document Server patches

This directory is the **canonical home** of Euro-Office modifications to ONLYOFFICE / Euro-Office Document Server.

Do not keep production patch logic only on a server, in an agent worktree, or in an old MCP repository.

## Layout

```text
documentserver-patches/
  README.md
  <document-server-version>/
    <patch-family>/
      README.md
      manifest.json
      patcher / source patch
      optional tests and evidence pointers
```

Current imported patch:

- `9.3.4.60/chart-public-getters/` — historical production patch imported from
  `joskapista44/euro-mcp-public:feature/documentserver-getter-patch`.

The full generated/minified Document Server bundle is intentionally not committed. A patch family records the exact
known input/output identities and deterministic transformations needed to reproduce the accepted state.

## Policy

1. **PUBLIC-FIRST:** if a new Document Server version provides the required semantics through a supported public API,
   prefer the native public API and retire the corresponding patch capability after LIVE acceptance.
2. **One patch maintenance boundary:** all required non-public DS integration is centralized here. MCP product code is
   being migrated away from direct ONLYOFFICE internal-model calls.
3. **Fail closed:** a patch for one upstream version must never be applied blindly to another version.
4. **Exact identity:** known original and accepted patched bundle hashes are version-specific evidence.
5. **Semantic acceptance:** matching method names/signatures is not proof that a capability still means the same thing.
6. **No silent carry-forward:** every Document Server upgrade runs the procedure in
   [docs/DOCUMENT-SERVER-PATCHES.md](../docs/DOCUMENT-SERVER-PATCHES.md).
7. **No deletion on upgrade:** old version directories remain historical/reproducibility evidence.
8. **No secrets or full customer documents** belong in this directory.

## Bridge direction

The target architecture is documented in
`docs/DS-BRIDGE-INTERNAL-MIGRATION-20261006.md`.

Long-term boundary:

```text
Euro-Office MCP
  ├─ supported public ONLYOFFICE API
  └─ Euro-Office DS Bridge
       └─ version-specific implementation in the patched Document Server
```

Exit target: zero direct ONLYOFFICE internal-model calls in MCP product code.
