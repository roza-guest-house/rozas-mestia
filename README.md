# Roza's Guest House website

The public website is published from the `main` branch through the existing GitHub Pages build and deployment workflow. Its address is https://roza-guest-house.github.io/rozas-mestia/.

## Basic website statistics

Cloudflare Web Analytics provides basic page-visit and performance statistics on the home and booking pages. It does not measure confirmed reservations, payments or advertising conversions. Google Analytics remains disabled.

The configuration is in `assets/js/basic-statistics-config.js`; the guarded loader is in `assets/js/basic-statistics.js`. The installation token is a public beacon identifier, not an account credential. Statistics are restricted to the explicit production origin and page paths. A future domain change requires updating this configuration and the Cloudflare site entry, then verifying the release.

The privacy page explains the collection and offers statistics preferences. It does not load the beacon. Browser Do Not Track or Global Privacy Control, a saved opt-out, or unavailable preference storage prevent loading. Preference changes apply to future page loads. Keep these controls and the privacy disclosure aligned with any future changes.

To turn off basic statistics for everyone, set `enabled` to `false` in `assets/js/basic-statistics-config.js` and publish through the same workflow. Check the actual public files after the deployment succeeds; a commit on GitHub alone does not prove the website has updated. Browser privacy settings and blockers can reduce the reported totals.

## Room photographs

Use the existing category folders as the owner's confirmed photograph mapping: `Classic`, `Peak Mountain View`, `Peak Mountain View with Balcony`, and `Cottage`. Preserve their arrangements when maintaining the galleries on both `index.html` and `book.html`. The hotel bathroom originals `IMG_9118.PNG` and `IMG_9122.PNG` occur in all three hotel categories. The cottage gallery uses its own photographs, including the separate `IMG_1560_result.jpeg` bathroom source. Keep the established room-type bathroom descriptions and avoid inventing additional layout or access claims from photographs.

## Release recovery

Before a release, retain a recoverable copy of the current published source and record its commit. Revert the relevant release commits through normal Git history if rollback is needed, allow the existing Pages workflow to complete, and verify the public files. Preserve the existing public address and booking backend configuration.
