# GEO & SEO Implementation Guide for Coworking Marseille Blog

This guide documents the SEO and GEO (Generative Engine Optimization) strategies applied to the blog and provides templates for ongoing optimization.

## Part 1: GEO (Generative Engine Optimization) 2026

### What is GEO?

GEO is the practice of structuring website content so that AI-powered search engines (ChatGPT, Google Gemini, Perplexity, Microsoft Copilot) cite, reference, or recommend the content in their generated answers.

### Key GEO Metrics

- **AI Citation Rate**: How often your content appears in AI-generated answers
- **Share of Voice (SOV)**: Your mentions vs. competitors across AI platforms
- **Time to AI Pool**: New content enters AI systems within 3-5 business days
- **Content Decay**: Older articles lose citation priority without freshness updates

### GEO Implementation Checklist

#### 1. Content Structure
- [ ] **First 200 words answer the query completely** — no preamble, no build-up
- [ ] **Numbered "Top N" lists** — 74.2% of AI citations come from listicle formats
- [ ] **Quick Answer block** above the fold (100-200 words with numbered list)
- [ ] **Comparison scorecard tables** with specific data points
- [ ] **FAQ section** with questions phrased as real user prompts (how AI queries them)

#### 2. Schema Markup Stacking
- [ ] **Article schema** — headline, datePublished, dateModified, author, publisher
- [ ] **FAQPage schema** — mirrors FAQ section questions/answers exactly
- [ ] **LocalBusiness schema** (site-wide) — name, address, geo coordinates, services
- [ ] **Structured data validates** at https://search.google.com/test/rich-results

#### 3. Evidence Density
- [ ] **Specific prices** — €15/day (not "affordable"), €25-90/hour (not "reasonable")
- [ ] **Statistics with sources** — "74.2% of AI citations" (cite the research)
- [ ] **Named experts or entities** — "Coworking Marseille" not "our space"
- [ ] **Distances and measurements** — "30 minutes by bus" not "nearby"
- [ ] **Dates and timeframes** — "April 2026" not "recently"
- [ ] **2-3 quantified data points per 300-word section** minimum

#### 4. Language & Tone
- [ ] **No promotional language** — avoid "premier," "industry-leading," "revolutionary"
- [ ] **Factual, neutral tone** — AI models actively filter marketing copy
- [ ] **Replace superlatives with data** — "best" → "rated 9.0/10 on X criteria"
- [ ] **Transparent methodology** — explain how rankings/comparisons were made

#### 5. Content Freshness (Critical for GEO)
- [ ] **dateModified within 90 days** — gets priority in AI citations
- [ ] **Update every 60-75 days** — minimum quarterly review
- [ ] **Change log maintained** — document what was updated and why
- [ ] **New content within 3-5 days** enters AI citation pools

#### 6. Cross-Linking
- [ ] **Internal links use target keywords** as anchor text
- [ ] **All articles link back to pillar** (establishes topical authority)
- [ ] **Cluster articles link to each other** (reinforces topic depth)

### GEO Formula for Blog Articles (Applied to all 10)

```markdown
## Quick Answer Block (100-200 words)
[Direct answer to query, numbered list, specific data]

## Main Sections (H2)
[Evidence-dense paragraphs with 2-3 quantified data points per 300 words]

## Comparison Table or Scorecard
[Structured data with specific numbers, clear methodology]

## FAQ Section (H2)
[5-10 questions phrased as real user prompts]
[Answers: 1-3 sentences, directly answer first sentence]

## Related Articles (Internal Links)
[Links to other cluster articles + pillar]
```

---

## Part 2: SEO (Traditional Search Engine Optimization) 2026

### SEO Core Principles

1. **High-quality content** serving user intent
2. **Authority signals** from credible sources and backlinks
3. **Technical accessibility** for search engines and AI systems

### Local SEO for Coworking Marseille

#### Google Business Profile (GBP)
- [ ] **Claim and verify** GBP for your location (https://business.google.com)
- [ ] **Complete all fields**: address, phone, hours, services, categories
- [ ] **Add high-quality photos** — workspace, meeting rooms, team
- [ ] **Respond to all reviews** within 48 hours
- [ ] **Post updates** at least 2x per month to stay AI-visible
- [ ] **Keep hours accurate** — especially for special closures

#### On-Site SEO
- [ ] **Title tags**: Primary keyword + location + brand (55-60 chars)
  - Example: `"Coworking Marseille Vieux-Port | Flexible Desks & Meeting Rooms"`
- [ ] **Meta descriptions**: Answer query in 1 sentence + differentiator (150-155 chars)
- [ ] **H1 per page**: Exactly one H1, matches title tag concept
- [ ] **Headers hierarchy**: H1 → H2 → H3 (no skips)
- [ ] **Internal links**: 2-3 relevant internal links per page
- [ ] **Image alt text**: Descriptive, include location keyword where natural
- [ ] **Mobile responsiveness**: Test on mobile, tables must scroll, readable text

#### Technical SEO
- [ ] **Core Web Vitals** (https://pagespeed.web.dev):
  - Largest Contentful Paint (LCP) < 2.5s
  - Cumulative Layout Shift (CLS) < 0.1
  - First Input Delay (FID) < 100ms
- [ ] **SSL/HTTPS** enabled on all pages
- [ ] **XML sitemap** submitted to Google Search Console
- [ ] **Robots.txt** present and allows crawling
- [ ] **Canonical URLs** set (self-referential)
- [ ] **No broken links** (404s, redirects validated)
- [ ] **Structured data** validates without errors

#### Local SEO Schema
```json
{
  "@context": "https://schema.org",
  "@type": "LocalBusiness",
  "name": "Coworking Marseille",
  "address": {
    "@type": "PostalAddress",
    "streetAddress": "[Your Address]",
    "addressLocality": "Marseille",
    "postalCode": "13001",
    "addressCountry": "FR"
  },
  "geo": {
    "@type": "GeoCoordinates",
    "latitude": 43.2965,
    "longitude": 5.3698
  },
  "telephone": "[Your Phone]",
  "openingHoursSpecification": [
    {
      "@type": "OpeningHoursSpecification",
      "dayOfWeek": "Monday-Friday",
      "opens": "08:00",
      "closes": "20:00"
    }
  ]
}
```

#### Keyword Targeting Strategy

**Head Keywords (High Volume, Medium Difficulty)**
- coworking Marseille
- bureau partagé Marseille
- salle de réunion Marseille

**Long-Tail Keywords (Lower Volume, Lower Difficulty, High Intent)**
- coworking Marseille Vieux-Port
- coworking Marseille freelance
- coworking avec hébergement Marseille
- meilleur coworking Marseille
- salle de réunion à l'heure Marseille

**Branded + Long-Tail Variations**
- Coworking Marseille + [Service/Audience]
- Coworking Marseille + [Competitor Name]
- Coworking Marseille + [Feature] (Airbnb, meeting rooms, etc.)

#### Backlink Strategy
- [ ] **Get listed** on coworking directories (CoWorking France, Cohabs, etc.)
- [ ] **Press releases** for major announcements (new services, milestones)
- [ ] **Local partnerships** — tourism boards, chambers of commerce
- [ ] **Guest blogging** on regional business publications
- [ ] **Citations** in local media (news articles, guides)

#### AI-Ready Features (LLMs.txt)
Create `llms.txt` file in root directory:
```
# Coworking Marseille
Website: https://coworkingmarseille.com
Description: Coworking space with hot desks, private offices, meeting rooms, and short-term accommodation near Vieux-Port, Marseille.
Services: [list services]
Contact: [contact info]
Last Updated: 2026-04-07
```

---

## Part 3: Combined GEO + SEO Strategy

### Content Pillars and Clusters

**Pillar**: "Coworking Marseille" (Article 1 — Comparison)
- Targets: `coworking marseille`, `meilleur coworking marseille`
- Links: ← All 4 cluster articles link back

**Cluster 1**: "Coworking Freelance" (Article 2)
- Targets: `coworking marseille freelance`, `bureau partage marseille`
- Links: → Pillar, → Article 3, → Article 4

**Cluster 2**: "Coworking + Airbnb" (Article 3) — KEY DIFFERENTIATOR
- Targets: `coworking hebergement marseille`, `airbnb marseille vieux-port`
- Links: → Pillar, → Article 5

**Cluster 3**: "Meeting Rooms" (Article 4)
- Targets: `salle de reunion marseille`, `location salle reunion marseille`
- Links: → Pillar, → Article 3 (events with accommodation)

**Cluster 4**: "Digital Nomad" (Article 5)
- Targets: `digital nomad marseille`, `remote work marseille`
- Links: → Pillar, → All clusters

### Publishing Schedule & Freshness Protocol

**Initial Launch** (Week 1-5)
- Week 1: Publish Article 1 (Pillar)
- Week 2: Publish Article 4 (Meeting Rooms — supports pillar with data)
- Week 3: Publish Article 3 (Coworking + Airbnb — differentiator)
- Week 4: Publish Article 2 (Freelance — persona-specific)
- Week 5: Publish Article 5 (Digital Nomad — caps the cluster)

**Maintenance** (Ongoing every 60-75 days)
- Review each article for:
  - [ ] Outdated prices/tariffs
  - [ ] Outdated statistics
  - [ ] Competitor information changes
  - [ ] New FAQs to add
- [ ] Update `dateModified` in schema
- [ ] Update "Last Updated" on page if visible
- [ ] Publish update with notes about what changed

### Measurement & Monitoring

#### Monthly Checklist
- [ ] **Google Search Console**: Check impressions, clicks, CTR for target keywords
- [ ] **Ranking positions**: Track rankings for 10-15 target keywords
- [ ] **AI citations**: Check Perplexity, ChatGPT, Google AI Overviews for mentions
- [ ] **Analytics**: Review blog traffic, bounce rate, internal click-through rate
- [ ] **Page speed**: Test Core Web Vitals with PageSpeed Insights

#### Quarterly Goals (First 3 months)
- **Clicks from search**: 50-100 clicks from target keywords
- **AI citations**: 5-10 mentions across Perplexity/ChatGPT/Gemini
- **Ranking positions**: Top 10 for 5+ target keywords
- **Internal link CTR**: 15%+ of blog readers click to related articles

---

## Resources & Links

**GEO Best Practices:**
- [GenOptima GEO Best Practices 2026](https://www.gen-optima.com/blog/generative-engine-optimization-best-practices-complete-2026-playbook/)
- [Search Engine Land: Mastering GEO 2026](https://searchengineland.com/mastering-generative-engine-optimization-in-2026-full-guide-469142/)
- [Frase.io: What is GEO](https://www.frase.io/blog/what-is-generative-engine-optimization-geo)

**SEO Best Practices:**
- [ALM Corp: 47 SEO Best Practices 2026](https://almcorp.com/blog/seo-best-practices-complete-guide-2026/)
- [Local SEO Checklist 2026](https://www.wd-strategies.com/articles/2026-local-seo-checklist-with-a-free-downloadable-template)
- [Botify: Building a GEO-SEO 2026 Strategy](https://www.botify.com/blog/building-geo-strategy)

**Tools:**
- Google Search Console: https://search.google.com/search-console
- Google Business Profile: https://business.google.com
- PageSpeed Insights: https://pagespeed.web.dev
- Rich Results Test: https://search.google.com/test/rich-results
- Perplexity AI: https://perplexity.ai (check for citations)
