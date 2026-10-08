# Contributing

## Setup

Use Node.js 22 or 24 LTS and npm. Clone the repository, then run:

```sh
npm ci
npm run verify
```

Keep runtime code under src/, package tooling under scripts/, source smoke tests
under src/__tests__/, and installed-package consumers under test/. The build is
generated into dist/ and is intentionally not committed. Commit package-lock.json
when changing dependencies. Do not commit credentials, personal spreadsheets, or
unrelated generated files.

## Checks

- `npm test`: source API, worker, and stream regression tests.
- `npm run build`: all four outputs, workers, source maps, and declarations.
- `npm run build:min`: rebuild minified variants only.
- `npm run test:types`: type consumers; run build first.
- `npm run test:package`: build, pack, install outside the repo, and test consumers.
- `npm run pack:check`: build and inspect the tarball allowlist.
- `npm run audit:dependencies`: reject high/critical dependency advisories.
- `npm run verify`: required release checks, including dependency audit.
- `npm run test:perf`: optional 50,000-row file round-trip benchmark.

Verification requires registry access to audit dependencies and install the
isolated consumer; it does not publish. Scripts use npm's own CLI path so Windows
does not require POSIX shell commands. Invoke tooling through npm scripts.

## Pull Requests

Keep changes focused and add coverage for affected API contracts. Check both
inline and worker behavior when changing serialization or routing. Worker scripts
must remain deployable in every output format. Keep declarations and the README
consistent with runtime behavior. Do not add browser support or pooling claims
without implementing and testing them.

## Compatibility and Support

Node 22/24 on Linux and Windows are the CI support matrix. Newer supported Node
releases are added after validation; the engine floor does not promise that every
future major release has already been tested. Node 20 and earlier are unsupported.
ESM/CommonJS root and minified exports share one public API. Internal src/dist
paths are not public exports.

Before 1.0, breaking API changes require a minor version bump and migration notes;
patches are intended to preserve compatibility. After 1.0, follow SemVer rules.
Only the latest published minor line receives routine fixes; older versions are
supported on a best-effort basis. This is a community project with no response-time
SLA. Report reproducible non-security issues through GitHub issues; use private
reporting for security issues.