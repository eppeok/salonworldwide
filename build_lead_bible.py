from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

doc = Document()

# Set default font
style = doc.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(11)

# Helper: Add styled heading
def add_heading(text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x1B, 0x3A, 0x5C)
    return h

# Helper: Add styled table
def add_table(headers, rows):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Header row
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = h
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(10)
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="2E75B6"/>')
        cell._tc.get_or_add_tcPr().append(shading)
    # Data rows
    for row_data in rows:
        row = table.add_row()
        for i, val in enumerate(row_data):
            row.cells[i].text = str(val)
            for p in row.cells[i].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(10)
    return table

# Helper: Add bullet list
def add_bullets(items):
    for item in items:
        p = doc.add_paragraph(item, style='List Bullet')
        for run in p.runs:
            run.font.size = Pt(10)

def add_bold_para(bold_text, normal_text=""):
    p = doc.add_paragraph()
    run = p.add_run(bold_text)
    run.bold = True
    run.font.size = Pt(11)
    if normal_text:
        run2 = p.add_run(normal_text)
        run2.font.size = Pt(11)
    return p

# ===================== COVER PAGE =====================
for _ in range(6):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("LEAD RESEARCH BIBLE")
run.bold = True
run.font.size = Pt(36)
run.font.color.rgb = RGBColor(0x1B, 0x3A, 0x5C)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Salon Marketing Masterclass — US Outreach Campaign")
run.font.size = Pt(16)
run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("EvolvXAI  |  evolvxai.com/salon01-3329")
run.font.size = Pt(12)
run.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)

doc.add_page_break()

# ===================== SECTION 1: MARKET OVERVIEW =====================
add_heading("SECTION 1: US SALON MARKET OVERVIEW", 1)

doc.add_paragraph("The US salon industry is massive, fragmented, and growing — making it ideal for targeted outreach with a $97 masterclass offer.")

add_heading("Market Size", 2)
add_bullets([
    "Total salons in the US: 1,077,381 enterprises (2026, IBISWorld)",
    "Hair salon market: $60.0 billion (2026)",
    "Hair + nail combined: $90.9 billion (2025)",
    "Projected growth: $95.99 billion by 2033 (CAGR 6.78%)",
    "Industry growth: 5.5% CAGR (2020-2025)",
    "80% of beauty salons are independently owned and operated",
    "60%+ of salon owners are female",
    "BLS projects +5% growth (2024-2034) for hairstylists and cosmetologists"
])

add_heading("Top 10 States by Salon Density", 2)
add_table(
    ["Rank", "State", "Est. Salons", "Key Cities", "Notes"],
    [
        ["1", "Florida", "14,787+", "Miami, Orlando, Tampa, Jacksonville", "Highest density. Tourism + retirees + lifestyle."],
        ["2", "California", "13,616+", "LA, SF, San Diego, Sacramento", "Premium services, multicultural, eco-focused."],
        ["3", "Texas", "12,000+", "Houston, Dallas, Austin, San Antonio", "Fastest growing. Men's grooming rising."],
        ["4", "New York", "11,000+", "NYC, Brooklyn, Long Island, Buffalo", "Luxury market, high fashion culture."],
        ["5", "Illinois", "8,000+", "Chicago, Naperville, Aurora", "Strong independent salon scene."],
        ["6", "Pennsylvania", "7,500+", "Philadelphia, Pittsburgh, Allentown", "Urban/suburban mix."],
        ["7", "Georgia", "6,500+", "Atlanta, Savannah, Augusta", "Growing metro areas."],
        ["8", "Ohio", "6,000+", "Columbus, Cleveland, Cincinnati", "Midwest hub."],
        ["9", "North Carolina", "5,500+", "Charlotte, Raleigh, Durham", "Fast-growing southeast."],
        ["10", "New Jersey", "5,000+", "Newark, Jersey City, Hoboken", "High density suburban."],
    ]
)

add_heading("Target Audience Profile", 2)
add_table(
    ["Attribute", "Target", "Rationale"],
    [
        ["Location", "United States — all states, focus on FL, CA, TX, NY", "Maximum salon density"],
        ["Business type", "Salon owners: hair, beauty, nail, blowout bars", "Matches course content"],
        ["Pain signal", "Running Facebook/Instagram ads with poor ROI", "Ready to switch to organic"],
        ["Revenue stage", "$10k-$100k/month revenue", "Can afford $97, motivated by growth"],
        ["Platforms active", "Instagram, Facebook, TikTok, Google", "Where we find and reach them"],
        ["Owner profile", "Independent owner, 1-50 employees, female-majority", "Decision-maker, can buy immediately"],
    ]
)

doc.add_page_break()

# ===================== SECTION 2: LEAD SOURCE LAYERS =====================
add_heading("SECTION 2: LEAD SOURCE LAYERS (7 Layers)", 1)
doc.add_paragraph("Each layer represents a distinct source of US salon owner leads. Work layers 1-3 simultaneously in Week 1, then add layers 4-7 from Week 2 onward.")

# LAYER 1
add_heading("LAYER 1: Google Maps + Yelp (Direct Scraping)", 2)
doc.add_paragraph("Google Maps contains 200M+ businesses. Scraping salon listings gives you name, phone, email, website, reviews — the richest cold outreach data available.")

add_bold_para("Search Queries to Run:")
add_table(
    ["Search Query", "Target Cities", "Expected Volume"],
    [
        ['"hair salon" + [city name]', "Miami, LA, Houston, NYC, Chicago, Dallas, Atlanta", "500-1,000 per city"],
        ['"beauty salon" + [city name]', "Same + Phoenix, San Diego, Tampa, Orlando", "300-500 per city"],
        ['"nail salon" + [city name]', "Same cities", "200-400 per city"],
        ['"salon suite" + [city name]', "Same cities", "100-200 per city"],
        ['"blowout bar" + [city name]', "NYC, LA, Miami, Chicago", "50-100 per city"],
    ]
)

add_bold_para("Tools:")
add_bullets([
    "G Maps Extractor (Chrome extension) — one-click export to CSV from Google Maps search results",
    "Outscraper — bulk scraping across multiple geographies at once",
    "Leads Sniper — extracts fresh leads with emails, phone, social profiles",
    "MapLeadScraper — no-code, works on Google Maps + Yelp + TripAdvisor",
    "Yelp Web Scraper Tool (Chrome extension) — extracts business data from Yelp listings",
])

add_bold_para("Data You Get: ", "Business name, owner name, phone, email, website, address, ratings, review count, social media profiles, business category")

add_bold_para("Qualifying Filters:")
add_bullets([
    "3+ star rating (active, not failing business)",
    "Has recent reviews (last 30 days = active business)",
    "Has a website (more established, higher revenue)",
    "Has email or phone listed (reachable)",
    "Located in target city/state",
])

# LAYER 2
add_heading("LAYER 2: Apollo.io / Hunter.io (B2B Email Prospecting)", 2)
doc.add_paragraph("Apollo.io has 210M+ contacts with 65+ data attributes. This is your primary source for verified salon owner emails at scale.")

add_bold_para("Apollo.io Exact Search Filters:")
add_table(
    ["Filter", "Value", "Notes"],
    [
        ["Industry", '"Beauty" OR "Cosmetics" OR "Personal Care"', "LinkedIn-aligned categories"],
        ["Job Title", '"Owner" OR "Founder" OR "Salon Owner" OR "CEO" OR "Managing Director"', "Decision-makers only"],
        ["SIC Code", "7231 (Beauty Shops)", "US government classification"],
        ["NAICS Code", "812111 (Barber Shops), 812112 (Beauty Salons)", "More specific classification"],
        ["Location", "United States (priority: FL, CA, TX, NY)", "Filter by state for targeted outreach"],
        ["Company Size", "1-50 employees", "Small/medium salons"],
        ["Revenue", "$100K - $10M", "Established businesses"],
        ["Email Status", "Verified only", "91% accuracy rate"],
    ]
)

add_bold_para("Hunter.io Approach:")
add_bullets([
    "Use salon websites found via Google Maps scraping",
    "Run domain search to find owner email addresses",
    "Verify all emails before adding to GHL",
    "Expected yield: 500-1,000 verified emails per week",
])

add_bold_para("Cost:")
add_bullets([
    "Apollo.io Free: 100 credits/month (1 credit = 1 email reveal)",
    "Apollo.io Basic: $49/month for 5,000 credits",
    "Hunter.io Free: 25 searches/month",
    "Hunter.io Starter: $49/month for 500 searches",
])

# LAYER 3
add_heading("LAYER 3: Instagram (DM Outreach)", 2)
doc.add_paragraph("Instagram is the highest-ROI cold channel for US salon owners. DMs feel personal, get read, and start real conversations.")

add_bold_para("Hashtags to Search (by tier):")
add_table(
    ["Tier", "Hashtags", "Purpose"],
    [
        ["TIER 1\n(Owner-focused)", "#salonowner #salonowners #salonownerlife #salonbusiness\n#salonbusinessowner #salonownership #beautypreneur #saloncoach", "Highest signal — these people self-identify as owners"],
        ["TIER 2\n(Business)", "#salonmarketing #salongrowth #salonownersuccess\n#beautybusiness #beautyindustry #salonentrepreneur", "Business-minded salon professionals"],
        ["TIER 3\n(General)", "#salonlife #hairsalon #beautysalon #modernsalon\n#americansalon #behindthechair", "Broader pool — filter carefully for owners"],
        ["TIER 4\n(Location)", "#miamisalon #lasalon #nysalon #houstonsalon\n#atlantasalon #chicagosalon #dallassalon #tampasalon\n#orlandosalon #phoenixsalon #sandiegosalon", "City-specific — highest targeting precision"],
    ]
)

add_bold_para("Bio Keywords That Indicate OWNER (not just stylist):")
add_bullets([
    '"salon owner" / "owner of" / "founder" / "my salon" / "our salon"',
    '"CEO" / "beauty entrepreneur" / "salon suite owner"',
    '"boss babe" / "beauty boss" / "girl boss" (informal signals)',
    'Booking link in bio (Vagaro, Square, Booksy, Fresha, GlossGenius)',
])

add_bold_para("Qualifying Signals:")
add_bullets([
    "Follower count: 500-10,000 (small business owner, not influencer)",
    "Posts business content (not just hair photos — sharing tips, behind-the-scenes)",
    "Has booking link in bio (established business)",
    "Located in US (check bio location tag)",
    "Signs of running ads: boosted posts, 'Sponsored' label, Meta Business Suite",
    "Posts regularly (last post within 7 days = active)",
])

add_bold_para("Volume: ", "30-50 DMs/day MANUALLY. Do NOT use automation bots — Instagram will restrict your account.")

# LAYER 4
add_heading("LAYER 4: Facebook Groups (Organic Presence)", 2)

add_table(
    ["Group Name", "Focus", "Est. Members", "Strategy"],
    [
        ["Salon Owners Group", "General business", "Large", "Comment value Week 1, post Week 2+"],
        ["Hair Salon Owners Network", "Hair-specific networking", "Large", "Share tips, answer questions"],
        ["Salon Owners United", "Community support", "Medium", "Engage daily, build rapport"],
        ["The Salon Owners Social Media Marketing Group", "Social media marketing", "Medium", "MOST RELEVANT — share marketing insights"],
        ["Help for Hairstylists", "Employee salons, NJ-based", "27k+", "Target owner members specifically"],
        ["Salon Owners Exclusive Networking Forum", "Private networking", "Medium", "Build relationships, no hard sell"],
        ["Ask Mags", "Florida community", "Medium", "Good for FL market targeting"],
        ["PocketSuite Community", "Booking/tech users", "Medium", "Tech-savvy salon owners"],
    ]
)

add_bold_para("Rules:")
add_bullets([
    "Week 1: ONLY comment on others' posts. Give genuine value. Zero promotion.",
    "Week 2+: One organic post per group per week. NO links in post body ever.",
    "Soft CTA only: 'DM me if you want details' or 'drop a comment'",
    "Post schedule: Monday/Wednesday/Friday rotation across groups",
    "Read and follow each group's rules — some ban any form of selling",
])

# LAYER 5
add_heading("LAYER 5: LinkedIn (Professional Outreach)", 2)

add_bold_para("LinkedIn Sales Navigator Filters:")
add_table(
    ["Filter", "Value"],
    [
        ["Title", '"Salon Owner" OR "Beauty Salon Owner" OR "Hair Salon Owner" OR "Founder"'],
        ["Industry", '"Cosmetics" OR "Health, Wellness and Fitness" OR "Consumer Services"'],
        ["Geography", "United States"],
        ["Company headcount", "1-50"],
        ["Posted on LinkedIn", "Past 30 days (active users only)"],
    ]
)

add_bold_para("Boolean Search Strings (Regular LinkedIn):")
add_bullets([
    '"salon owner" AND ("United States" OR "USA")',
    '"beauty salon" AND "founder" AND (California OR Texas OR Florida OR "New York")',
    '"hair salon" AND "owner" AND (Miami OR Houston OR "Los Angeles" OR Chicago)',
])

add_bold_para("Hashtags to Follow/Engage: ", "#salonowner #beautyindustry #salonbusiness #beautybusiness #salonmarketing #beautypreneur")

add_bold_para("Volume: ", "20-30 connection requests/day with personalized notes. Keep notes under 300 characters.")

# LAYER 6
add_heading("LAYER 6: Reddit (Community Engagement)", 2)
add_table(
    ["Subreddit", "Focus", "Strategy"],
    [
        ["r/Hairstylist", "Professional stylists discussing business", "Answer business questions, build credibility"],
        ["r/Barbers", "Barbers and barbering business", "Cross-sell to barber/salon hybrids"],
        ["r/Hair", "General hair community", "Light presence, link to value content"],
        ["r/HaircareScience", "Science-based hair care", "Position as expert, not salesperson"],
        ["r/FancyFollicles", "Hair showcasing", "Engage with professionals in comments"],
    ]
)
add_bold_para("Strategy: ", "Long-game only. Provide genuine value for 2-3 weeks before any mention of course. Build karma. Never spam links. This is a credibility channel, not a sales channel.")

# LAYER 7
add_heading("LAYER 7: Threads (Emerging Channel)", 2)
add_bullets([
    "Search salon-related keywords and follow salon owners posting business content",
    "Mirror Instagram strategy but text-focused — share insights, start conversations",
    "Lower competition than Instagram — earlier mover advantage",
    "Reply to salon owner posts with genuine value before any pitch",
    "Lighter touch — use as supplementary channel to Instagram",
])

doc.add_page_break()

# ===================== SECTION 3: LEAD SCORING =====================
add_heading("SECTION 3: LEAD SCORING FRAMEWORK", 1)
doc.add_paragraph("Score every lead 0-100 to prioritize follow-up. Update scores as leads engage across channels.")

add_table(
    ["Signal", "Points", "Rationale"],
    [
        ["Located in US (confirmed)", "+10", "Required — non-US leads are disqualified"],
        ["Confirmed salon owner (not employee)", "+10", "Must be the decision-maker"],
        ["Currently running paid ads", "+15", "Highest pain point match for our offer"],
        ["Has 500-10K followers on Instagram", "+5", "Right-size business for this offer"],
        ["Active on social media (posted in last 7 days)", "+5", "Reachable and engaged"],
        ["Has booking system in bio (Vagaro, Square, etc.)", "+5", "Established business with systems"],
        ["Estimated revenue $10K-$100K/month", "+5", "Can afford $97, motivated by growth"],
        ["Replied to DM or email", "+20", "Engaged — highest intent signal"],
        ["Clicked landing page link", "+15", "Showing real purchase interest"],
        ["Multiple touchpoints (engaged on 2+ channels)", "+10", "Warmer across the board"],
    ]
)

add_heading("Scoring Tiers", 2)
add_table(
    ["Tier", "Score Range", "Action", "Follow-Up Cadence"],
    [
        ["HOT LEAD", "70-100", "Priority follow-up, WhatsApp voice note, personal DM", "Daily until converted or declined"],
        ["WARM LEAD", "40-69", "Nurture email sequence, continued DM engagement", "Every 2-3 days"],
        ["COOL LEAD", "20-39", "Keep in email sequence, occasional value DM", "Weekly"],
        ["COLD", "0-19", "Park for later batch. Focus effort elsewhere.", "Monthly or end-of-campaign push"],
    ]
)

doc.add_page_break()

# ===================== SECTION 4: TOP 50 US CITIES =====================
add_heading("SECTION 4: TOP 50 US CITIES TARGET LIST", 1)
doc.add_paragraph("Prioritized by salon density, population, and market activity. Start with Tier 1, expand to Tier 2 in Week 2, and Tier 3 in Week 3+.")

cities = [
    # Tier 1
    ("Miami", "FL", "1", "Highest salon density state. Tourism + lifestyle."),
    ("Los Angeles", "CA", "1", "Massive market. Premium services."),
    ("Houston", "TX", "1", "Fastest growing. Diverse salon market."),
    ("New York City", "NY", "1", "Luxury market. Highest spend."),
    ("Chicago", "IL", "1", "Strong independent salon scene."),
    ("Dallas", "TX", "1", "Large metro. Rising demand."),
    ("Atlanta", "GA", "1", "Growing metro. Beauty hub."),
    ("Tampa", "FL", "1", "High salon density. Warm climate market."),
    ("Orlando", "FL", "1", "Tourism + local demand."),
    ("Phoenix", "AZ", "1", "Fast-growing southwest."),
    # Tier 2
    ("San Diego", "CA", "2", "Premium coastal market."),
    ("San Antonio", "TX", "2", "Large metro, growing."),
    ("Philadelphia", "PA", "2", "Urban/suburban mix."),
    ("San Francisco", "CA", "2", "High-end market."),
    ("Charlotte", "NC", "2", "Fast-growing southeast."),
    ("Jacksonville", "FL", "2", "FL market expansion."),
    ("Austin", "TX", "2", "Trendy, young market."),
    ("Nashville", "TN", "2", "Booming entertainment city."),
    ("Denver", "CO", "2", "Growing metro."),
    ("Las Vegas", "NV", "2", "Tourism-driven beauty market."),
    ("Sacramento", "CA", "2", "CA expansion market."),
    ("Fort Worth", "TX", "2", "DFW metro expansion."),
    ("Raleigh", "NC", "2", "Research Triangle growth."),
    ("Seattle", "WA", "2", "Tech-savvy salon owners."),
    ("Portland", "OR", "2", "Indie salon culture."),
    ("Minneapolis", "MN", "2", "Midwest hub."),
    ("Columbus", "OH", "2", "Growing midwest."),
    ("Indianapolis", "IN", "2", "Affordable market."),
    ("Virginia Beach", "VA", "2", "Coastal VA market."),
    ("Brooklyn", "NY", "2", "Separate from NYC — indie salons."),
    # Tier 3
    ("Cleveland", "OH", "3", "Midwest expansion."),
    ("St. Louis", "MO", "3", "Central US hub."),
    ("Kansas City", "MO", "3", "Growing metro."),
    ("Baltimore", "MD", "3", "East coast expansion."),
    ("Pittsburgh", "PA", "3", "PA expansion market."),
    ("Cincinnati", "OH", "3", "Midwest expansion."),
    ("Oklahoma City", "OK", "3", "Southern expansion."),
    ("Louisville", "KY", "3", "Growing market."),
    ("Memphis", "TN", "3", "Southern beauty market."),
    ("New Orleans", "LA", "3", "Unique cultural market."),
    ("Salt Lake City", "UT", "3", "Growing western market."),
    ("Richmond", "VA", "3", "East coast expansion."),
    ("Birmingham", "AL", "3", "Southern market."),
    ("Milwaukee", "WI", "3", "Midwest market."),
    ("Albuquerque", "NM", "3", "Southwest expansion."),
    ("Tucson", "AZ", "3", "AZ expansion."),
    ("Omaha", "NE", "3", "Central US."),
    ("Tulsa", "OK", "3", "Oklahoma expansion."),
    ("Fresno", "CA", "3", "CA inland market."),
    ("Honolulu", "HI", "3", "Island market — premium."),
]

add_table(
    ["City", "State", "Tier", "Notes"],
    [(c[0], c[1], c[2], c[3]) for c in cities]
)

doc.add_page_break()

# ===================== SECTION 5: PLATFORM CHECKLISTS =====================
add_heading("SECTION 5: PLATFORM SETUP CHECKLISTS", 1)

add_heading("GoHighLevel (GHL) Setup", 2)
add_bullets([
    "Create or access GHL sub-account for this campaign",
    "Set up custom email sending domain (secondary domain, NOT main EvolvXAI domain)",
    "Warm email domain for minimum 2 weeks before sending cold emails",
    "Import lead lists as contacts (CSV upload from Apollo.io / Google Maps scraping)",
    "Build email workflow: 9 emails across 3 phases with appropriate delays",
    "Set up pipeline stages: Cold > Contacted > Replied > Interested > Link Clicked > Purchased",
    "Configure tags: lead_source, lead_tier, campaign_phase",
    "Set up landing page tracking (if GHL landing page) or UTM parameters for evolvxai.com/salon01-3329",
    "Test all email sequences with internal test contacts before going live",
])

add_heading("Instagram Setup", 2)
add_bullets([
    "Switch to Professional/Business account if not already",
    "Optimize bio: mention salon marketing, organic growth, clear CTA",
    "Add evolvxai.com/salon01-3329 as link in bio",
    "Create Story Highlights: Testimonials, Results, About, Masterclass",
    "Prepare first 3-5 feed posts (value content about salon marketing)",
    "Save hashtag sets for quick use during DM prospecting",
    "Set up Quick Replies for common DM responses",
])

add_heading("Facebook Setup", 2)
add_bullets([
    "Ensure personal profile looks credible (professional photo, relevant bio)",
    "Join all 8 target groups from Layer 4 list",
    "Set notification preferences for groups (to see new posts quickly)",
    "Prepare comment templates for common discussion topics",
    "Do NOT post anything promotional in Week 1 — comment only",
])

add_heading("LinkedIn Setup", 2)
add_bullets([
    "Update headline to mention salon marketing / helping salon owners",
    "Optimize About section with relevant positioning",
    "Prepare 3 connection request note templates (under 300 chars each)",
    "Save Sales Navigator searches (if using premium)",
    "Follow salon industry hashtags",
])

add_heading("Apollo.io / Hunter.io Setup", 2)
add_bullets([
    "Create account and select appropriate plan",
    "Save search filters: Industry + Job Title + SIC Code + Location",
    "Build first list: 500 verified US salon owner emails",
    "Export to CSV for GHL import",
    "Set up email verification workflow",
])

add_heading("Google Maps / Yelp Scraping Setup", 2)
add_bullets([
    "Install G Maps Extractor Chrome extension",
    "Install Yelp Web Scraper Chrome extension (optional)",
    "Run first test scrape: 'hair salon Miami' on Google Maps",
    "Verify data quality: check email/phone accuracy on first 10 results",
    "Set up scraping schedule: 2 cities per day during Week 1",
])

doc.add_page_break()

# ===================== SECTION 6: WEEKLY KPI TARGETS =====================
add_heading("SECTION 6: WEEKLY KPI TARGETS", 1)

add_table(
    ["Metric", "Week 1-2\n(Attention)", "Week 3-4\n(Interest)", "Week 5\n(Desire)", "Week 6\n(Action)"],
    [
        ["IG DMs sent/day", "30-50", "20-30", "15-20", "10 (re-engage)"],
        ["Cold emails sent/day (GHL)", "150-200", "50 nurture", "50 last-chance", "50 close"],
        ["LinkedIn requests/day", "20-30", "15-20", "10-15", "5-10"],
        ["Facebook group posts/week", "0 (comment only)", "3-5", "3-5", "2-3"],
        ["WhatsApp messages/day", "0", "0", "5-10 (warm leads)", "10-20 (all leads)"],
        ["Landing page visits", "50-100", "200-400", "400+", "400+ (retarget)"],
        ["Email open rate", "—", "35-45%", "40-55%", "40-55%"],
        ["Email CTR", "—", "5-8%", "8-12%", "10-15%"],
        ["Conversions (sales)", "0-5 early", "10-20", "15-25", "15-30"],
        ["Total revenue target", "$0-$485", "$970-$1,940", "$1,455-$2,425", "$1,455-$2,910"],
    ]
)

add_bold_para("Sales Target: ", "40-60 units in 6 weeks = $3,880-$5,820 at $97")
add_bold_para("Realistic Upside: ", "80-120 units = $7,760-$11,640")

doc.add_page_break()

# ===================== SECTION 7: INDUSTRY RESOURCES =====================
add_heading("SECTION 7: PROFESSIONAL ASSOCIATIONS & INDUSTRY RESOURCES", 1)

add_table(
    ["Resource", "URL", "Value for This Campaign"],
    [
        ["Professional Beauty Association (PBA)", "probeauty.org", "Largest beauty org — industry trends, networking events"],
        ["International SalonSpa Business Network (ISBN)", "salonspanetwork.org", "Multi-unit salon executives — high-value targets"],
        ["Associated Hair Professionals (AHP)", "associatedhairprofessionals.com", "Insurance + business support — active community"],
        ["Salon Today", "salontoday.com", "Business ideas + trends — content inspiration"],
        ["Modern Salon", "modernsalon.com", "Industry news — stay relevant in conversations"],
        ["American Salon", "americansalonmag.com", "Industry coverage — credibility building"],
        ["Openmart US Beauty Salon Database", "openmart.com/data/beauty-salons", "Dedicated salon lead database with verified data"],
    ]
)

add_heading("Key Conferences & Events (Networking Opportunities)", 2)
add_bullets([
    "ISBN Annual Conference 2026 — Colorado Springs (executive salon leaders)",
    "IBS New York (International Beauty Show) — largest US beauty trade show",
    "Premiere Beauty Shows — Orlando, Columbus (regional events)",
    "PBA Beauty Week — annual industry gathering",
])

doc.add_page_break()

# ===================== SECTION 8: CRITICAL REMINDERS =====================
add_heading("CRITICAL CAMPAIGN REMINDERS", 1)

add_bold_para("1. You have ZERO US audience right now. ")
doc.add_paragraph("Phase 1 must be aggressive — 50+ touchpoints per day across DMs and cold email. Do not dilute effort across all 7 channels simultaneously in Week 1. Start with Instagram DMs and cold email ONLY. Layer in other channels from Week 2 onward.")

add_bold_para("2. Channel Priority Order:")
add_bullets([
    "1 — Instagram DM outreach (fastest warm-up, highest response rate)",
    "2 — Cold email via GHL (scalable, trackable, automatable)",
    "3 — Facebook group organic presence (trust-building, compounds over time)",
    "4 — LinkedIn outreach (targets multi-location owners, professional channel)",
    "5 — Short-form content (slow burn, builds authority over time)",
    "6 — WhatsApp/SMS direct follow-up (highest close rate, Phase 3 only)",
    "7 — Threads (supplementary, emerging channel)",
])

add_bold_para("3. Weekly Review Cadence:")
add_bullets([
    "Every Monday: Review DM reply rate, email open rate, and landing page traffic",
    "Every Wednesday: Check Facebook group engagement and adjust tone if needed",
    "Every Friday: Review email CTR, retargeting ad frequency and CPM",
    "End of Phase 2: Count warm leads (replies + opt-ins) — if under 50, add more cold outreach",
    "End of Phase 3: Count conversions — if under 10 sales, extend deadline by 5 days",
])

# SAVE
path = "/Users/nishamore/Documents/Salon world wide/Lead_Research_Bible.docx"
doc.save(path)
print(f"Lead Research Bible saved to: {path}")
print(f"File size: {os.path.getsize(path)} bytes")
