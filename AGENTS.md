# AGENTS.md — personal

Personal projects hub for Anatide ANANI.

## Structure

- `portfolio/` — Portfolio site (French). **Complete, deployed on GitHub Pages.** Vanilla HTML/CSS/JS. Google Fonts (Bebas Neue, DM Mono, Plus Jakarta Sans). Custom cursor, dark/light theme (localStorage), scroll reveal via IntersectionObserver, responsive.
- `curriculum-vitae/` — Online CV (French). Printable A4 layout, same design system as portfolio.

## Notes

- **No build step** — open `index.html` in a browser.
- **No package.json, no framework, no bundler** — pure vanilla frontend for both projects.
- Google Fonts loaded via `<link>` in the `<head>`, not via npm.
- Theme preference persists in `localStorage` key `"theme"`.
- French language throughout (content, labels, READMEs).
- `.gitignore` only ignores `*.DS_Store`.
