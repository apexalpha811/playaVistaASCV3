/**
 * Seed the CMS with Playa Vista ASC content.
 *
 * Usage:
 *   1. npm install @supabase/supabase-js (or use the one from cms-dashboard)
 *   2. Fill in SITE_ID below (copy from Supabase after adding the site in the dashboard)
 *   3. node seed-cms.mjs
 */

import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = process.env.SUPABASE_URL || 'https://vcmptzcnftrwoiolibqi.supabase.co'
const SUPABASE_SERVICE_KEY = process.env.SUPABASE_SERVICE_KEY // set in .env.local
const SITE_ID = process.env.SITE_ID || '5357dcff-efb2-42da-a35c-160fa6e5cf22'

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_KEY)

const pages = [
  {
    slug: 'overview',
    title: 'Overview',
    content: {
      hero_eyebrow: 'Established 2026 · Playa Vista, CA',
      hero_claim_1: 'We are a hybrid surgery center and wellness clinic, built from day one to keep adapting to surgical innovations.',
      hero_claim_2: 'The OR power, space, and data infrastructure are spec\'d up front to absorb surgical automation and advanced recovery as each proves out. The facility modernizes by design, not by costly retrofit.',
      hero_claim_3: 'The clinic is projected to be cash-flow positive on its own, more than 10 months before the first OR opens.',
      problem_eyebrow: 'The Problem',
      mission_body: 'We bring outpatient surgery, regenerative medicine, and longevity care under one roof. That means one record, one continuous relationship, and better outcomes than three disconnected providers can deliver.',
      vision_body: 'Surgery is modernizing as automation, data, and recovery science reshape the operating room. We are building to move with that curve instead of retrofitting behind it.',
      philosophy_body: 'We build the bones ahead of the technology and let operating cash flow fund each upgrade. The facility evolves with the field rather than aging against it.',
    },
    seo: {
      title: 'Playa Vista ASC & Rejuvenation Clinic — Investor Pitch',
      description: 'A hybrid surgery center and wellness clinic in Playa Vista, CA. Raising $5M via SAFE. Phase 1 clinic cash-flow positive at Month 3–6.',
    },
  },
  {
    slug: 'opportunity',
    title: 'Opportunity',
    content: {
      opp_unserved: 'No luxury ASC on the Westside integrates wellness into surgical care.',
      opp_dual_stream: 'Insurance-backed surgical revenue and cash-pay wellness revenue, combined under one roof.',
      opp_cash_flow: 'Phase 1 clinic revenue starts more than 10 months before OR activation. Capital works while accreditation runs.',
      forward_lead: 'The forward posture is not a promise. It is in the buildout spec.',
      forward_infra: 'The OR power, layout, and data backbone are specced to accept surgical automation later, without a rebuild. This raise funds the bones, not the future hardware.',
      forward_throughput: 'Surgical automation compresses recovery and standardizes case times. That drives more predictable scheduling and more cases per room over the life of the facility.',
      forward_cashflow: 'Next-generation surgical platforms come later, paid for by operating cash flow and equipment financing rather than dilutive equity in this round.',
    },
    seo: {},
  },
  {
    slug: 'services',
    title: 'Services',
    content: {
      services_lead: '6,000–6,500 SF in Playa Vista. Phase 1 launches at Month 3–6; Phase 2 ORs activate at Month 16.',
    },
    seo: {},
  },
  {
    slug: 'model',
    title: 'Model & Financials',
    content: {},
    seo: {},
  },
  {
    slug: 'facility',
    title: 'Facility & Timeline',
    content: {},
    seo: {},
  },
  {
    slug: 'investment',
    title: 'Investment',
    content: {
      ask_sub: 'To open the Phase 1 clinic, build out two ORs and a procedure room, and carry working capital through Month 16 license activation.',
      raise_lead: 'We are raising a SAFE released in two milestone-tied tranches, so capital deploys against demonstrated progress and the larger tranche prices after the clinic proves the model.',
    },
    seo: {},
  },
  {
    slug: 'team',
    title: 'Team & Next Steps',
    content: {
      cta_heading: "LET'S BUILD THE WESTSIDE'S FIRST INTEGRATED ASC",
      cta_sub: 'We deliver vertically integrated surgical care and longevity under one roof in Playa Vista.',
    },
    seo: {},
  },
  {
    slug: 'appendix',
    title: 'Appendix',
    content: {},
    seo: {},
  },
]

async function seed() {
  if (SITE_ID === 'PASTE_SITE_ID_HERE') {
    console.error('Set SITE_ID before running. Copy it from Supabase after adding the site in the CMS dashboard.')
    process.exit(1)
  }

  console.log('Seeding', pages.length, 'pages for site', SITE_ID)

  for (const page of pages) {
    const { error } = await supabase.from('pages').upsert(
      { site_id: SITE_ID, ...page },
      { onConflict: 'site_id,slug' }
    )
    if (error) {
      console.error('Failed:', page.slug, error.message)
    } else {
      console.log('Seeded:', page.slug)
    }
  }

  console.log('Done.')
}

seed()
