# Releasing ExcelIO

## One-Time Registry Setup

- Verify ownership of npm scope `@dipanshuhandoo` and availability of the package
  name. An unavailable or unauthorized scope/name blocks release.
- Confirm package.json repository, bugs, homepage, author, license, description,
  and Node support fields. Enable private vulnerability reporting on GitHub.
- Enable npm account 2FA and authenticate with `npm login` using npm's supported
  browser/terminal workflow. Verify identity with `npm whoami`. Never route
  passwords, OTPs, or tokens through chat, code, or committed configuration.
- If npm requires a publishing token, use a current granular token with package
  write access and an appropriate expiry. Keep it in the approved credential
  mechanism, not in this repository. CI currently validates only.
- Configure branch protection requiring all Validate Package matrix jobs.

## Release Procedure

1. Review a clean release checkout and passing Node 22/24 Linux/Windows CI.
2. Review current dependency advisories, including the moderate finding documented
   in SECURITY.md. Resolve or explicitly accept/defer it; a passing high-severity
   audit gate is not a declaration that there are no vulnerabilities.
3. Update CHANGELOG.md with the release date and version. For the first release,
   use the existing 0.1.0 version. For later releases, explicitly select patch,
   minor, or major according to the compatibility policy.
4. Run the commands below and inspect the packed file list and npm dry-run output.
5. Only after review/approval, run the public publication command.
6. Install the published version in a fresh project and verify inline and worker
   read/write operations. Record the matching Git tag and GitHub release after
   successful publication; do not include credentials in release output.

```sh
npm ci
npm run verify
npm run pack:check
npm run publish:dry-run
```

Explicit publication:

```sh
npm run release
```

`publish:public` is an alias for direct public publishing. `release` does not bump
versions, make commits, or create tags. Both use the public npm registry and run
prepublishOnly verification followed by prepack rebuilding. Authentication/2FA
may require direct user interaction. Dry-run checks packing and lifecycle scripts,
but does not prove account permissions, name availability, or successful publishing.

For later versions, `npm version patch --no-git-tag-version` (or minor/major) updates
package.json and the lockfile without an automatic commit/tag. Review and commit
those changes explicitly before publishing. The initial setup never publishes.

## Package Gates

The tarball must contain package.json, README, LICENSE, CHANGELOG, and the exact
dist artifacts for readable/minified ESM/CommonJS. Each build includes both worker
bundles, source maps, and matching declarations. Source files, tests, plans, build
scripts, temporary files, and secrets are excluded. Installed consumers, not just
source tests, must pass. Internal packing uses --ignore-scripts to avoid recursive
prepack execution; normal npm packing runs the build.

Full verification and isolated installs require registry access. High/critical
audit findings stop verification. Do not bypass failed checks to ship a release.

## Recovery

Published npm versions cannot be overwritten. If publication fails, inspect npm's
error and verify registry state before retrying the same version. After a bad
release, issue a corrected new version and consider `npm deprecate` with a migration
message. Do not treat unpublishing as a dependable rollback strategy. Investigate
credential exposure immediately and revoke affected credentials through npm.