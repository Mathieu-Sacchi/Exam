# Blog Publishing Guide for coworkingmarseille.com

## Overview

This blog contains 5 bilingual (FR + EN) articles optimized for SEO and GEO (Generative Engine Optimization). Each article includes embedded JSON-LD schemas and meta tag recommendations.

## File Structure

```
blog/
├── fr/                                          # French articles
│   ├── 01-meilleur-coworking-marseille-2026.html      (PILLAR - Comparison)
│   ├── 02-coworking-freelance-marseille-guide-2026.html (Freelance Guide)
│   ├── 03-coworking-hebergement-marseille-vieux-port.html (Airbnb + Coworking)
│   ├── 04-salle-reunion-marseille-location-tarifs.html  (Meeting Rooms)
│   └── 05-digital-nomad-marseille-guide-2026.html       (Digital Nomad Guide)
├── en/                                          # English articles
│   ├── 01-best-coworking-marseille-2026.html
│   ├── 02-coworking-freelance-marseille-guide-2026.html
│   ├── 03-coworking-accommodation-marseille-vieux-port.html
│   ├── 04-meeting-room-marseille-rental-prices.html
│   └── 05-digital-nomad-marseille-guide-2026.html
├── schema/
│   └── localbusiness.json                       # Site-wide LocalBusiness schema
└── PUBLISHING-GUIDE.md                          # This file
```

## Publishing on Wix

### Step 1: Create Blog Posts

For each article (FR first, then EN):

1. Go to **Wix Dashboard > Blog > Create New Post**
2. Copy the article content (everything inside `<article>` tags)
3. Paste into the Wix blog editor
4. Format headings (H1, H2, H3), tables, and lists using the Wix editor tools
5. Set the **URL slug** as specified in the HTML comment at the top of each file

### Step 2: Configure SEO Settings (per post)

In the Wix blog post editor, click **SEO Settings**:

1. **Title tag**: Copy from the HTML comment at the top of each file
2. **Meta description**: Copy from the HTML comment at the top of each file
3. **URL slug**: Set as specified (e.g., `meilleur-coworking-marseille-2026`)
4. **OG Image**: Upload a custom 1200x630px image with readable text overlay

### Step 3: Add JSON-LD Structured Data (per post)

Each article contains 2 JSON-LD blocks (`<script type="application/ld+json">`):

1. Go to **SEO Settings > Advanced SEO > Structured Data Markup**
2. Click **Add New Markup**
3. Paste the **Article schema** JSON-LD block
4. Click **Add New Markup** again
5. Paste the **FAQPage schema** JSON-LD block

### Step 4: Add Site-Wide LocalBusiness Schema

1. Go to **Wix Dashboard > Settings > Custom Code** (or **Wix Velo** if using Base44)
2. Click **Add Custom Code**
3. Paste the content from `schema/localbusiness.json` wrapped in `<script type="application/ld+json">` tags
4. Set placement to **Head** and apply to **All pages**
5. **IMPORTANT**: Replace placeholder values:
   - `[VOTRE ADRESSE]` → Your actual street address
   - `+33-XXXXXXXXX` → Your actual phone number
   - Update social media URLs
   - Verify GPS coordinates (43.2965, 5.3698) match your location

### Step 5: Set Up Bilingual Pages

Option A — Wix Multilingual (recommended):
1. Enable **Wix Multilingual** in Dashboard > Settings
2. Add English as a secondary language
3. Create English versions of each post using the EN files
4. Wix will automatically add hreflang tags

Option B — Manual subdirectory:
1. Create English blog posts with `/en/blog/` prefix in URLs
2. Manually add hreflang tags in each page's custom code:
```html
<link rel="alternate" hreflang="fr" href="https://coworkingmarseille.com/blog/[fr-slug]" />
<link rel="alternate" hreflang="en" href="https://coworkingmarseille.com/en/blog/[en-slug]" />
```

## Publishing Order

| Week | Article | Why this order |
|------|---------|----------------|
| 1 | Article 1 (Pillar - Comparison) | Establishes topical authority, all others link back |
| 2 | Article 4 (Meeting Rooms) | Supports pillar with service-specific data |
| 3 | Article 3 (Coworking + Airbnb) | Key differentiator, references pillar + Art 4 |
| 4 | Article 2 (Freelance Guide) | Persona-specific, links to all prior articles |
| 5 | Article 5 (Digital Nomad Guide) | Caps the cluster, international audience |

## URL Mapping

| Article | French URL | English URL |
|---------|-----------|-------------|
| 1 (Pillar) | `/blog/meilleur-coworking-marseille-2026` | `/en/blog/best-coworking-marseille-2026` |
| 2 (Freelance) | `/blog/coworking-freelance-marseille-guide-2026` | `/en/blog/coworking-freelance-marseille-guide-2026` |
| 3 (Airbnb) | `/blog/coworking-hebergement-marseille-vieux-port` | `/en/blog/coworking-accommodation-marseille-vieux-port` |
| 4 (Meeting) | `/blog/salle-reunion-marseille-location-tarifs` | `/en/blog/meeting-room-marseille-rental-prices` |
| 5 (Nomad) | `/blog/digital-nomad-marseille-guide-2026` | `/en/blog/digital-nomad-marseille-guide-2026` |

## Content Freshness (GEO requirement)

Pages with `dateModified` within the past 90 days get priority in AI search results.

**Every 60-75 days**, do the following for each article:
1. Update at least one price, statistic, or piece of factual information
2. Optionally add a new FAQ question
3. Update `dateModified` in the Article JSON-LD schema
4. Republish the post on Wix (which updates the "Last modified" date)

## Validation Checklist (after publishing each article)

- [ ] Article is live and accessible at the correct URL
- [ ] Title tag and meta description are correct (check in browser tab + page source)
- [ ] JSON-LD schemas validate at https://search.google.com/test/rich-results
- [ ] Internal links to other blog articles work
- [ ] Internal links to service pages work
- [ ] Images load correctly and have alt text
- [ ] Mobile layout is readable (tables, lists, text)
- [ ] hreflang tags point to the correct FR/EN counterpart
- [ ] OG Image displays correctly when sharing on social media (test at https://opengraph.dev)

## Monitoring

- **Google Search Console**: Submit sitemap, monitor impressions/clicks for target keywords
- **AI Citation Monitoring**: Monthly check on Perplexity, ChatGPT, Google AI Overviews for queries like "coworking Marseille", "best coworking Marseille", "coworking with accommodation Marseille"
- **Schema Validation**: Re-test rich results after any content update
- **Analytics**: Track blog page views, time on page, and conversions (booking clicks, contact form submissions)
