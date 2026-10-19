# 08 — Source Research: PlacePlate (@placeplate12345)

> This file documents **who the source is, everything that was found, and how it was
> researched** (Oct 2026). The other files in this repo turn these findings into
> reusable generation instructions. All techniques are rewritten in our own words;
> nothing here reproduces PlacePlate's paid files or long-form copyrighted text.

## The creator

| Item | Detail |
|---|---|
| Brand | **PlacePlate** — PowerPoint template shop + custom presentation design service |
| Instagram | [@placeplate12345](https://www.instagram.com/placeplate12345/) — "Place Plate", ~111K followers, 191 posts, 0 following |
| IG bio | "PlacePlate PowerPoint Template 👇👇 and PowerPoint Service" (link-in-bio → Linktree) |
| Specialty | **Morph-transition PowerPoint templates** — every product is built around PowerPoint's Morph feature |
| Social proof claimed | "Trusted by 3,000+ clients & customers"; testimonials from Germany, Egypt, France, Belgium, Mexico, Saudi Arabia, Zambia, … |
| Based in | Indonesia (IDR storefront, Tokopedia, Indonesian-language checkout; sells globally in USD) |
| Contact | helo@placeplate.com (support replies "within 24 hours"), "Book Call" CTA on site |

## Every link found

**Profiles / hubs**

- Instagram — https://www.instagram.com/placeplate12345/
- TikTok — https://www.tiktok.com/@placeplate (same videos cross-posted from IG Reels)
- Linktree — https://linktr.ee/PlacePlate ("Please, select a location to continue" → service page / US store / Indonesian store / other countries)
- Website — https://placeplate.com/ (templates, service, blog, FAQ, about)
- YouTube — channel "PlacePlates", https://www.youtube.com/channel/UCi-PBqZzStANSmfGTT_CxeQ (~595 subs; 2 preview videos embedded per template product page)
- Pinterest — https://pinterest.com/PlacePlate321
- X — https://x.com/place_plate543
- Facebook — https://www.facebook.com/share/1XcLVdv7DA

**Stores / checkout**

- International (USD): Lemon Squeezy storefront `placeplates.lemonsqueezy.com`; site footer also references Gumroad `placeplate.gumroad.com`
- Indonesia: `placeplate.shop` + Tokopedia `tokopedia.com/placeplate`
- Also accepts crypto (USDT) — payment details deliberately **not** copied into this repo.

**Key site pages mined for this skill**

- Catalog: https://placeplate.com/templates/ (47 listings — fully indexed in `07-template-catalog.md`)
- Service/pricing: https://placeplate.com/service/ (Basic $35/10 slides layout-only; Professional $75/10 slides incl. Morph & smart animation, optional intro logo animation; process = order form → pay → design → revisions/delivery)
- FAQ: https://placeplate.com/faq/ (100% editable; PowerPoint 2019+ incl. Mac; instant download; custom files delivered via Drive/WeTransfer/email as .pptx)
- Blog (11 articles, all read): Morph guide, pitch deck guide, startup pitch deck, academic presentation guide, 10 slide-design tips, data-visualization guide, 7 common mistakes, how to choose a template, best business templates, PowerPoint vs Google Slides, free templates list.

## What their Reels/TikToks actually are (content formula)

Instagram blocks unauthenticated scraping (HTTP 403), so individual Reels could not be
watched frame-by-frame. However, **the same videos are cross-posted to their TikTok**,
whose captions/descriptions are indexed. Combined with the YouTube preview videos
embedded on every product page, the reel formula is fully reconstructable:

1. **Template showcase (≈15 s)** — a screen recording flip-through of one template's
   slides so the Morph motion *is* the content. Caption pattern:
   `Morph PowerPoint Template ✨🔥 #powerpoint #presentation #morphtransition #powerpointpresentation #powerpointtemplate`.
   CTA is always the link-in-bio (Linktree → store).
2. **Step-by-step Morph tutorials** — teach the duplicate-slide → move objects →
   apply Morph workflow (their most-saved format; one tutorial indexed at ~363K views).
3. **Themed showcases** — e.g. a "Netflix-style" Morph template video (~215K views,
   ~16K likes) proving that recognizable UI/theme skins + Morph travel well.
4. **Service proof** — anonymized excerpts of client decks (also on the service page,
   with a confidentiality note).

Recurring hashtags: `#powerpoint #morphtransition #powerpointpresentation
#powerpointtemplate #presentation #tutorial`. Audio is usually their own original
sound. Posting is high-volume (191 IG posts) and product-led: each reel maps to a
SKU in the catalog.

**Lesson for generation:** their entire content engine = *one deck, presented as motion*.
Any deck we generate with real Morph can be re-purposed the same way: export the
slideshow as MP4 (File → Export → Create a Video), crop 9:16, caption with the
formula above.

## Their offer, in one paragraph

47 Morph-based .pptx products (PP-1…PP-49 + 7 bundles), 16:9, fully editable,
PowerPoint 2019+, Windows/Mac, personal + commercial license ("do not resell"),
instant download. Standard templates = 12 layouts (~$3.50); popular ones $5;
flagship PP-1 Pro Pitch Deck = 20 layouts, Master-Slide + smart color system,
free fonts Onest/Inter ($7); three free 5-slide samplers (PP-46/47/48). Bundles up
to the $79 All-in-One "lifetime access" pack. Positioning line: **"Present with
Clarity. Design with Impact."** — modern design, consistency, storytelling,
credibility. Full index: `07-template-catalog.md`.

## Research method & honesty notes

- Read in full: homepage, catalog (all pages), service page, FAQ snippets, 8 of 11
  blog articles verbatim via page fetch, remainder via indexed excerpts; 3 product
  pages in depth (PP-1, PP-20, PP-46).
- Instagram profile metadata comes from search-index snippets (follower/post
  counts, bio). Reels themselves require login — **not** scraped; represented via
  the creator's own TikTok cross-posts + YouTube previews.
- Pinterest also returned 403; treat it as a distribution mirror of the catalog.
- Hex palettes in `02-design-system.md` / the generator are **reconstructions**
  matching each template's named color story and preview descriptions — their
  actual .pptx internals were not purchased or extracted.
- Prices/counts are a snapshot; check placeplate.com for current values.
- This repo is a *style & technique* skill inspired by PlacePlate's public teaching.
  It does not redistribute their templates; if you want the originals, buy them —
  they're inexpensive and this research links straight to each product.
