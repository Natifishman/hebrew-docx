# Publishing hebrew-docx

One-time setup:

```bash
npm login                      # creates/uses your npm account
```

Check the package name is free (it was at the time this was written):

```bash
npm view hebrew-docx
# should print: npm error 404 'hebrew-docx@latest' is not in this registry
```

Before every release:

```bash
npm install                    # installs the real 'docx' dependency
npm test                       # must pass
npm run example                # sanity-check the two example docs build
```

Publish:

```bash
npm publish --access public
```

Tag the release in git so the GitHub release matches the npm version:

```bash
git tag v0.1.0
git push origin v0.1.0
```

Then on GitHub: Releases → Draft a new release → pick the tag → publish. A
release makes the repo show up in GitHub's "Releases" feed, which is worth
having even for a small library.

## Bumping the version later

```bash
npm version patch   # or minor / major
npm publish --access public
git push --follow-tags
```

## Repo setup checklist (do once, on GitHub)

- [ ] Description field: "Generate Word documents in Hebrew/RTL from Node.js — correct bidirectional text, tables, and layout."
- [ ] Topics: `docx`, `word`, `hebrew`, `rtl`, `bidi`, `nodejs`, `document-generation`
- [ ] Add the npm version badge to the top of README once published:
      `[![npm](https://img.shields.io/npm/v/hebrew-docx)](https://www.npmjs.com/package/hebrew-docx)`
- [ ] Settings → General → uncheck "Wikis" and "Projects" if unused, keep "Issues" on
