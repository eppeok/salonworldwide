#!/usr/bin/env python3
"""
Generate Lead Research Bible DOCX for Salon Marketing Masterclass.
Run: python3 generate_lead_bible.py
Requires: pip install python-docx
"""

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def set_cell_shading(cell, color_hex):
    """Apply background shading to a table cell."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_table_borders(table):
    """Set borders on the entire table."""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '  <w:top w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '  <w:left w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '  <w:right w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '</w:tblBorders>'
    )
    tblPr.append(borders)


def add_styled_table(doc, headers, rows, header_color="1F3864", col_widths=None):
    """Create a table with styled header row."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    set_table_borders(table)

    # Header row
    hdr = table.rows[0]
    for i, text in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.name = "Calibri"
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, header_color)

    # Data rows
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        for c_idx, text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(text))
            run.font.size = Pt(10)
            run.font.name = "Calibri"
            # Alternate row shading
            if r_idx % 2 == 1:
                set_cell_shading(cell, "F2F2F2")

    # Column widths
    if col_widths:
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = Inches(width)

    return table


def add_bullet(doc, text, bold_prefix=None, level=0):
    """Add a bullet point paragraph."""
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.5 + level * 0.25)
    if bold_prefix:
        run_b = p.add_run(bold_prefix)
        run_b.bold = True
        run_b.font.size = Pt(11)
        run_b.font.name = "Calibri"
        run = p.add_run(text)
        run.font.size = Pt(11)
        run.font.name = "Calibri"
    else:
        run = p.add_run(text)
        run.font.size = Pt(11)
        run.font.name = "Calibri"
    return p


def add_body(doc, text):
    """Add a normal body paragraph."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(11)
    run.font.name = "Calibri"
    return p


def add_sub_heading(doc, text):
    """Add a bold sub-heading."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12)
    run.font.name = "Calibri"
    run.font.color.rgb = RGBColor(31, 56, 100)
    return p


# ---------------------------------------------------------------------------
# Document creation
# ---------------------------------------------------------------------------

doc = Document()

# -- Default font setup --
style = doc.styles["Normal"]
font = style.font
font.name = "Calibri"
font.size = Pt(11)

# Update heading styles
for i in range(1, 4):
    hs = doc.styles[f"Heading {i}"]
    hs.font.name = "Calibri"
    hs.font.color.rgb = RGBColor(31, 56, 100)

# =========================================================================
# COVER PAGE
# =========================================================================
for _ in range(6):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("LEAD RESEARCH BIBLE")
run.bold = True
run.font.size = Pt(36)
run.font.name = "Calibri"
run.font.color.rgb = RGBColor(31, 56, 100)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Salon Marketing Masterclass \u2014 US Outreach Campaign")
run.font.size = Pt(18)
run.font.name = "Calibri"
run.font.color.rgb = RGBColor(89, 89, 89)

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("EvolvXAI")
run.bold = True
run.font.size = Pt(16)
run.font.name = "Calibri"
run.font.color.rgb = RGBColor(31, 56, 100)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("evolvxai.com/salon01-3329")
run.font.size = Pt(14)
run.font.name = "Calibri"
run.font.color.rgb = RGBColor(0, 102, 204)
run.underline = True

for _ in range(4):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Product: Salon Marketing Masterclass ($97)")
run.font.size = Pt(12)
run.font.name = "Calibri"
run.font.color.rgb = RGBColor(89, 89, 89)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Target Market: US Salon Owners")
run.font.size = Pt(12)
run.font.name = "Calibri"
run.font.color.rgb = RGBColor(89, 89, 89)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Document Version 1.0 \u2014 2026")
run.font.size = Pt(11)
run.font.name = "Calibri"
run.font.color.rgb = RGBColor(140, 140, 140)

doc.add_page_break()

# =========================================================================
# TABLE OF CONTENTS (manual)
# =========================================================================
doc.add_heading("TABLE OF CONTENTS", level=1)
toc_items = [
    ("Section 1", "US Salon Market Overview"),
    ("Section 2", "Lead Source Layers (7 Layers)"),
    ("Section 3", "Lead Scoring Framework"),
    ("Section 4", "Top 50 US Cities Target List"),
    ("Section 5", "Platform-Specific Execution Checklists"),
    ("Section 6", "Weekly KPI Targets"),
    ("Section 7", "Professional Associations & Industry Resources"),
]
for num, title in toc_items:
    p = doc.add_paragraph()
    run = p.add_run(f"{num}: ")
    run.bold = True
    run.font.size = Pt(12)
    run.font.name = "Calibri"
    run = p.add_run(title)
    run.font.size = Pt(12)
    run.font.name = "Calibri"

doc.add_page_break()

# =========================================================================
# SECTION 1: US SALON MARKET OVERVIEW
# =========================================================================
doc.add_heading("SECTION 1: US SALON MARKET OVERVIEW", level=1)

add_sub_heading(doc, "Market Snapshot")
add_bullet(doc, "1,077,381 salon enterprises across the US (2026, IBISWorld)", bold_prefix="Total Salons: ")
add_bullet(doc, "$60 billion (hair salons); $90.9 billion (hair + nail salons combined)", bold_prefix="Market Size: ")
add_bullet(doc, "80% of salons are independently owned", bold_prefix="Ownership: ")
add_bullet(doc, "60%+ of salon owners are female", bold_prefix="Demographics: ")
add_bullet(doc, "5.5% CAGR (2020\u20132025)", bold_prefix="Growth Rate: ")

add_sub_heading(doc, "Top States by Salon Density")

state_headers = ["State", "Estimated Salons", "Key Cities", "Notes"]
state_rows = [
    ("Florida", "14,787+", "Miami, Orlando, Tampa, Jacksonville", "Highest salon density, tourism + retirees"),
    ("California", "13,616+", "LA, San Francisco, San Diego, Sacramento", "Premium services, multicultural, eco-focused"),
    ("Texas", "12,000+", "Houston, Dallas, Austin, San Antonio", "Fastest growing, men\u2019s grooming rising"),
    ("New York", "11,000+", "NYC, Brooklyn, Long Island, Buffalo", "Luxury market, high fashion culture"),
    ("Illinois", "8,000+", "Chicago, Naperville, Aurora", "Strong independent salon scene"),
    ("Pennsylvania", "7,500+", "Philadelphia, Pittsburgh, Allentown", "Mix of urban/suburban"),
    ("Georgia", "6,500+", "Atlanta, Savannah, Augusta", "Growing metro areas"),
    ("Ohio", "6,000+", "Columbus, Cleveland, Cincinnati", "Midwest hub"),
    ("North Carolina", "5,500+", "Charlotte, Raleigh, Durham", "Fast-growing southeast"),
    ("New Jersey", "5,000+", "Newark, Jersey City, Hoboken", "High density suburban"),
]
add_styled_table(doc, state_headers, state_rows, col_widths=[1.2, 1.3, 2.3, 2.0])

doc.add_page_break()

# =========================================================================
# SECTION 2: LEAD SOURCE LAYERS
# =========================================================================
doc.add_heading("SECTION 2: LEAD SOURCE LAYERS (7 LAYERS)", level=1)

# ----- LAYER 1 -----
doc.add_heading("LAYER 1: Google Maps + Yelp (Direct Scraping)", level=2)

add_sub_heading(doc, "What to Search")
add_bullet(doc, '"hair salon", "beauty salon", "nail salon", "blowout bar" + city name')

add_sub_heading(doc, "Recommended Tools")
add_bullet(doc, "G Maps Extractor (Chrome extension)")
add_bullet(doc, "Outscraper")
add_bullet(doc, "Leads Sniper")
add_bullet(doc, "MapLeadScraper")

add_sub_heading(doc, "Data Extracted")
add_bullet(doc, "Business name, owner name, phone, email, website, address, ratings, review count")

add_sub_heading(doc, "Target")
add_body(doc, "Top 50 US cities by salon density")

add_sub_heading(doc, "Priority Search Queries")
search_headers = ["Search Query", "Target Cities", "Expected Volume"]
search_rows = [
    ('"hair salon" + [city]', "Miami, LA, Houston, NYC, Chicago", "500\u20131,000 per city"),
    ('"beauty salon" + [city]', "Same + Atlanta, Phoenix, Dallas", "300\u2013500 per city"),
    ('"nail salon" + [city]', "Same", "200\u2013400 per city"),
    ('"salon suite" + [city]', "Same", "100\u2013200 per city"),
]
add_styled_table(doc, search_headers, search_rows, col_widths=[2.0, 2.5, 1.8])

add_sub_heading(doc, "Pro Tips")
add_bullet(doc, "Filter by 3+ star ratings")
add_bullet(doc, "Target active Google Business profiles")
add_bullet(doc, "Recent reviews = active business \u2014 prioritize these leads")

doc.add_page_break()

# ----- LAYER 2 -----
doc.add_heading("LAYER 2: Apollo.io / Hunter.io (B2B Email Prospecting)", level=2)

add_sub_heading(doc, "Apollo.io Exact Filters")
add_bullet(doc, '"Beauty" OR "Cosmetics" OR "Personal Care"', bold_prefix="Industry: ")
add_bullet(doc, '"Owner" OR "Founder" OR "Salon Owner" OR "CEO" OR "Managing Director"', bold_prefix="Job Title: ")
add_bullet(doc, "7231 (Beauty Shops)", bold_prefix="SIC Code: ")
add_bullet(doc, "812111 (Barber Shops), 812112 (Beauty Salons)", bold_prefix="NAICS Code: ")
add_bullet(doc, "United States (priority: FL, CA, TX, NY)", bold_prefix="Location: ")
add_bullet(doc, "1\u201350 employees", bold_prefix="Company Size: ")
add_bullet(doc, "$100K \u2013 $10M", bold_prefix="Revenue: ")

add_sub_heading(doc, "Hunter.io Approach")
add_body(doc, "Domain search for salon websites found via Google Maps scraping. Enter each salon\u2019s website domain to find verified email addresses associated with that domain.")

add_sub_heading(doc, "Expected Yield & Cost")
add_bullet(doc, "500\u20131,000 verified emails per week", bold_prefix="Expected Yield: ")
add_bullet(doc, "Apollo free tier = 100 credits/month; Basic = $49/month for 5,000 credits", bold_prefix="Cost: ")
add_bullet(doc, "Always verify before sending (91% accuracy on Apollo)", bold_prefix="Email Verification: ")

doc.add_page_break()

# ----- LAYER 3 -----
doc.add_heading("LAYER 3: Instagram (DM Outreach)", level=2)

add_sub_heading(doc, "Primary Hashtags to Search (Sorted by Relevance)")

add_body(doc, "").runs[0].bold = True
p = doc.add_paragraph()
run = p.add_run("TIER 1 \u2014 Owner-Focused:")
run.bold = True
run.font.size = Pt(11)
run.font.name = "Calibri"
run.font.color.rgb = RGBColor(31, 56, 100)
add_body(doc, "#salonowner  #salonowners  #salonownerlife  #salonbusiness  #salonbusinessowner  #salonownership  #beautypreneur  #saloncoach")

p = doc.add_paragraph()
run = p.add_run("TIER 2 \u2014 Business / Marketing:")
run.bold = True
run.font.size = Pt(11)
run.font.name = "Calibri"
run.font.color.rgb = RGBColor(31, 56, 100)
add_body(doc, "#salonmarketing  #salongrowth  #salonownersuccess  #beautybusiness  #beautyindustry  #salonentrepreneur")

p = doc.add_paragraph()
run = p.add_run("TIER 3 \u2014 General Salon:")
run.bold = True
run.font.size = Pt(11)
run.font.name = "Calibri"
run.font.color.rgb = RGBColor(31, 56, 100)
add_body(doc, "#salonlife  #hairsalon  #beautysalon  #modernsalon  #americansalon  #behindthechair")

p = doc.add_paragraph()
run = p.add_run("TIER 4 \u2014 Location-Specific:")
run.bold = True
run.font.size = Pt(11)
run.font.name = "Calibri"
run.font.color.rgb = RGBColor(31, 56, 100)
add_body(doc, "#miamisalon  #lasalon  #nysalon  #houstonsalon  #atlantasalon  #chicagosalon  #dallassalon  #tampasalon  #orlandosalon  #phoenixsalon")

add_sub_heading(doc, "Bio Keywords That Indicate Salon OWNER (Not Just Stylist)")
add_body(doc, '"salon owner", "owner of", "founder", "my salon", "our salon", "CEO", "beauty entrepreneur", "salon suite owner"')

add_sub_heading(doc, "Qualifying Signals")
add_bullet(doc, "500\u201310,000 (small business owner, not influencer)", bold_prefix="Follower Count: ")
add_bullet(doc, "Posts business content (not just hair photos)")
add_bullet(doc, "Has a booking link in bio (Vagaro, Square, Booksy, etc.)")
add_bullet(doc, "Located in US (check bio location)")
add_bullet(doc, "Signs of running ads: boosted posts, \u201cSponsored\u201d label")

add_sub_heading(doc, "Volume")
add_body(doc, "30\u201350 DMs per day (manual only, no automation bots).")

doc.add_page_break()

# ----- LAYER 4 -----
doc.add_heading("LAYER 4: Facebook Groups (Organic Presence)", level=2)

add_sub_heading(doc, "Complete Group List")

fb_headers = ["Group Name", "Focus", "Strategy"]
fb_rows = [
    ("Salon Owners Group", "General business", "Comment value Week 1, post Week 2+"),
    ("Hair Salon Owners Network", "Hair-specific", "Share tips, answer questions"),
    ("Salon Owners United", "Community support", "Engage daily, build rapport"),
    ("The Salon Owners Social Media Marketing Group", "Social media marketing", "MOST RELEVANT \u2014 share marketing insights"),
    ("Help for Hairstylists (27k+ members)", "Employee salons", "Target owner members specifically"),
    ("Salon Owners Exclusive Networking Forum", "Private networking", "Build relationships, no hard sell"),
    ("Ask Mags", "Florida-based community", "Good for FL market"),
    ("PocketSuite Community", "Booking/tech users", "Tech-savvy salon owners"),
]
add_styled_table(doc, fb_headers, fb_rows, col_widths=[2.5, 1.5, 2.8])

add_sub_heading(doc, "Rules of Engagement")
add_bullet(doc, "Week 1 = ONLY comment and help. Provide genuine value. Ask no one to buy anything.")
add_bullet(doc, "Week 2+ = One organic post per group per week maximum.")
add_bullet(doc, "NEVER drop links directly in posts. Links go in comments only if someone asks.")

add_sub_heading(doc, "Posting Schedule")
add_body(doc, "Monday / Wednesday / Friday rotation across groups.")

doc.add_page_break()

# ----- LAYER 5 -----
doc.add_heading("LAYER 5: LinkedIn (Professional Outreach)", level=2)

add_sub_heading(doc, "Sales Navigator Filters")
add_bullet(doc, '"Salon Owner" OR "Beauty Salon Owner" OR "Hair Salon Owner" OR "Founder" (+ industry filter)', bold_prefix="Title: ")
add_bullet(doc, '"Cosmetics" OR "Health, Wellness and Fitness" OR "Consumer Services"', bold_prefix="Industry: ")
add_bullet(doc, "United States", bold_prefix="Geography: ")
add_bullet(doc, "1\u201350", bold_prefix="Company Headcount: ")
add_bullet(doc, "Past 30 days (active users only)", bold_prefix="Posted on LinkedIn: ")

add_sub_heading(doc, "Boolean Search Strings (Regular LinkedIn)")
add_bullet(doc, '"salon owner" AND ("United States" OR "USA")')
add_bullet(doc, '"beauty salon" AND "founder" AND (California OR Texas OR Florida OR "New York")')
add_bullet(doc, '"hair salon" AND "owner" site:linkedin.com')

add_sub_heading(doc, "Hashtags to Follow & Engage")
add_body(doc, "#salonowner  #beautyindustry  #salonbusiness  #beautybusiness  #salonmarketing")

add_sub_heading(doc, "Volume")
add_body(doc, "20\u201330 connection requests per day with personalized notes.")

doc.add_page_break()

# ----- LAYER 6 -----
doc.add_heading("LAYER 6: Reddit (Community Engagement)", level=2)

add_sub_heading(doc, "Relevant Subreddits")
reddit_headers = ["Subreddit", "Description"]
reddit_rows = [
    ("r/Hairstylist", "Professional hairstylists discussing business"),
    ("r/Barbers", "Barbers and barbering business"),
    ("r/Hair", "General hair community"),
    ("r/FancyFollicles", "Hair showcasing"),
    ("r/HaircareScience", "Science-based hair care"),
]
add_styled_table(doc, reddit_headers, reddit_rows, col_widths=[2.0, 4.5])

add_sub_heading(doc, "Strategy")
add_bullet(doc, "Long-game only. Provide genuine value for 2\u20133 weeks before any mention of the course.")
add_bullet(doc, "Build karma through helpful comments and posts.")
add_bullet(doc, "NEVER spam links. Reddit communities will ban you instantly.")
add_bullet(doc, "Focus on building credibility as a marketing-savvy salon industry participant.")

doc.add_page_break()

# ----- LAYER 7 -----
doc.add_heading("LAYER 7: Threads (Emerging Channel)", level=2)

add_bullet(doc, "Search salon-related keywords on Threads.")
add_bullet(doc, "Follow and engage with salon owners posting business content.")
add_bullet(doc, "Lighter touch \u2014 emerging platform, lower competition.")
add_bullet(doc, "Mirror Instagram strategy but text-focused.")
add_bullet(doc, "Early mover advantage: less noise, easier to build organic reach.")

doc.add_page_break()

# =========================================================================
# SECTION 3: LEAD SCORING FRAMEWORK
# =========================================================================
doc.add_heading("SECTION 3: LEAD SCORING FRAMEWORK", level=1)

add_sub_heading(doc, "Scoring Criteria")

score_headers = ["Signal", "Points", "Rationale"]
score_rows = [
    ("Located in US (confirmed)", "+10", "Required \u2014 product is US-focused"),
    ("Is confirmed salon owner (not employee)", "+10", "Must be decision-maker"),
    ("Currently running paid ads", "+15", "Highest pain point match"),
    ("Has 500\u201310K followers on Instagram", "+5", "Right size business"),
    ("Active on social media (posts in last 7 days)", "+5", "Reachable"),
    ("Has booking system in bio", "+5", "Established business"),
    ("Salon revenue $10K\u2013$100K/month (estimated)", "+5", "Can afford $97"),
    ("Replied to DM or email", "+20", "Warm lead"),
    ("Clicked landing page", "+15", "Showing interest"),
    ("Multiple engagement touchpoints", "+10", "Warmer lead"),
]
add_styled_table(doc, score_headers, score_rows, col_widths=[3.0, 0.8, 2.8])

doc.add_paragraph()

add_sub_heading(doc, "Scoring Tiers")

tier_headers = ["Score Range", "Tier", "Action"]
tier_rows = [
    ("70\u2013100", "HOT LEAD", "Priority follow-up, WhatsApp voice note, direct call"),
    ("40\u201369", "WARM LEAD", "Nurture sequence, continue engagement across platforms"),
    ("20\u201339", "COOL LEAD", "Keep in email sequence, light touch, re-engage monthly"),
    ("0\u201319", "COLD", "Park for later, focus effort elsewhere"),
]
add_styled_table(doc, tier_headers, tier_rows, col_widths=[1.2, 1.3, 4.0])

doc.add_page_break()

# =========================================================================
# SECTION 4: TOP 50 US CITIES TARGET LIST
# =========================================================================
doc.add_heading("SECTION 4: TOP 50 US CITIES TARGET LIST", level=1)

city_headers = ["City", "State", "Est. Salons", "Priority Tier", "Notes"]
city_rows = [
    ("Miami", "FL", "3,500+", "Tier 1", "Highest density, diverse clientele"),
    ("Los Angeles", "CA", "4,000+", "Tier 1", "Premium market, celebrity culture"),
    ("Houston", "TX", "3,200+", "Tier 1", "Fastest-growing metro"),
    ("New York City", "NY", "5,000+", "Tier 1", "Largest market, luxury segment"),
    ("Chicago", "IL", "3,000+", "Tier 1", "Strong independent scene"),
    ("Dallas", "TX", "2,500+", "Tier 1", "High growth, affluent suburbs"),
    ("Atlanta", "GA", "2,800+", "Tier 1", "Major SE hub, diverse market"),
    ("Phoenix", "AZ", "2,200+", "Tier 1", "Rapid population growth"),
    ("San Diego", "CA", "1,800+", "Tier 1", "Affluent coastal market"),
    ("San Antonio", "TX", "1,600+", "Tier 1", "Growing metro area"),
    ("Philadelphia", "PA", "2,400+", "Tier 1", "Dense urban market"),
    ("Tampa", "FL", "1,500+", "Tier 1", "Tourism + residential mix"),
    ("Orlando", "FL", "1,400+", "Tier 1", "Tourism-driven economy"),
    ("Denver", "CO", "1,300+", "Tier 2", "Growing metro, health-conscious"),
    ("Charlotte", "NC", "1,200+", "Tier 2", "Fast-growing SE city"),
    ("Nashville", "TN", "1,100+", "Tier 2", "Booming economy, entertainment hub"),
    ("Las Vegas", "NV", "1,300+", "Tier 2", "Tourism + events driven"),
    ("Austin", "TX", "1,200+", "Tier 2", "Young, trendy, tech-savvy owners"),
    ("Jacksonville", "FL", "1,000+", "Tier 2", "Large metro area in FL"),
    ("San Francisco", "CA", "1,500+", "Tier 2", "Premium/luxury market"),
    ("Sacramento", "CA", "900+", "Tier 2", "Growing CA metro"),
    ("Seattle", "WA", "1,200+", "Tier 2", "Affluent, eco-conscious"),
    ("Portland", "OR", "900+", "Tier 2", "Independent, eco-focused"),
    ("Columbus", "OH", "1,000+", "Tier 2", "Midwest growth city"),
    ("Cleveland", "OH", "800+", "Tier 2", "Established market"),
    ("Raleigh", "NC", "700+", "Tier 2", "Research Triangle growth"),
    ("Minneapolis", "MN", "900+", "Tier 2", "Stable Midwest market"),
    ("St. Louis", "MO", "800+", "Tier 2", "Central US hub"),
    ("Kansas City", "MO", "700+", "Tier 2", "Growing metro"),
    ("Indianapolis", "IN", "800+", "Tier 2", "Midwest hub"),
    ("Milwaukee", "WI", "600+", "Tier 3", "Stable market"),
    ("Baltimore", "MD", "900+", "Tier 3", "Dense urban market"),
    ("Pittsburgh", "PA", "700+", "Tier 3", "Established market"),
    ("Cincinnati", "OH", "650+", "Tier 3", "Tri-state area"),
    ("Virginia Beach", "VA", "600+", "Tier 3", "Military + tourism market"),
    ("Oklahoma City", "OK", "500+", "Tier 3", "Growing market"),
    ("Louisville", "KY", "500+", "Tier 3", "Regional hub"),
    ("Memphis", "TN", "550+", "Tier 3", "Southern market"),
    ("Richmond", "VA", "500+", "Tier 3", "State capital market"),
    ("Birmingham", "AL", "450+", "Tier 3", "Largest AL metro"),
    ("New Orleans", "LA", "600+", "Tier 3", "Unique culture, tourism"),
    ("Salt Lake City", "UT", "500+", "Tier 3", "Growing Western market"),
    ("Albuquerque", "NM", "350+", "Tier 3", "Regional market"),
    ("Tucson", "AZ", "400+", "Tier 3", "Southern AZ market"),
    ("Boise", "ID", "300+", "Tier 3", "Fast-growing small metro"),
    ("El Paso", "TX", "350+", "Tier 3", "Border city market"),
    ("Honolulu", "HI", "400+", "Tier 3", "Island market, tourism"),
    ("Omaha", "NE", "350+", "Tier 3", "Stable Midwest market"),
    ("Tulsa", "OK", "400+", "Tier 3", "Regional OK market"),
    ("Fresno", "CA", "350+", "Tier 3", "Central Valley market"),
]
add_styled_table(doc, city_headers, city_rows, col_widths=[1.3, 0.5, 0.9, 0.9, 2.8])

doc.add_page_break()

# =========================================================================
# SECTION 5: PLATFORM-SPECIFIC EXECUTION CHECKLISTS
# =========================================================================
doc.add_heading("SECTION 5: PLATFORM-SPECIFIC EXECUTION CHECKLISTS", level=1)

# --- GHL ---
doc.add_heading("GoHighLevel (GHL) Setup Checklist", level=2)
ghl_items = [
    "Email domain purchased and connected",
    "Domain warming started (2\u20134 weeks before launch)",
    "Email sequences loaded (cold, nurture, last-chance, close)",
    "Pipeline stages configured (New Lead \u2192 Contacted \u2192 Engaged \u2192 Booked \u2192 Sold)",
    "Automations tested (trigger \u2192 email \u2192 wait \u2192 follow-up)",
    "Landing page connected and tracking pixels installed",
    "Webhook integrations tested (form fill \u2192 pipeline)",
    "SMS follow-up sequences configured (if applicable)",
    "Contact tags created: Source (IG, FB, Apollo, Google Maps), Status (Cold, Warm, Hot)",
]
for item in ghl_items:
    add_bullet(doc, item)

# --- Instagram ---
doc.add_heading("Instagram Checklist", level=2)
ig_items = [
    "Profile set to Professional/Business account",
    "Bio optimized: clear value proposition for salon owners",
    "Profile photo: professional headshot or brand logo",
    "Highlights ready: Testimonials, Results, About, FAQ",
    "10+ value posts published before outreach begins",
    "DM templates saved (3\u20135 variations)",
    "Hashtag research completed (Tier 1\u20134 saved)",
    "Target account list built (first 200 salon owners)",
]
for item in ig_items:
    add_bullet(doc, item)

# --- Facebook ---
doc.add_heading("Facebook Checklist", level=2)
fb_items = [
    "Personal profile looks credible (not a new/empty account)",
    "Profile photo, cover photo, and bio updated",
    "All target groups joined (allow 1\u20137 days for approval)",
    "Group rules read for each group",
    "Week 1 commenting schedule planned (5\u201310 comments/day across groups)",
    "Post templates drafted (value-first, no links)",
    "Engagement tracker spreadsheet created",
]
for item in fb_items:
    add_bullet(doc, item)

# --- LinkedIn ---
doc.add_heading("LinkedIn Checklist", level=2)
li_items = [
    "Headline updated to reflect salon marketing expertise",
    "About section optimized with relevant keywords",
    "Connection note templates ready (3\u20135 personalized variations)",
    "Sales Navigator filters saved (if using paid)",
    "Boolean search strings tested and saved",
    "Content plan: 2\u20133 posts/week about salon marketing insights",
    "Target list of 200+ salon owners built",
]
for item in li_items:
    add_bullet(doc, item)

# --- Apollo.io ---
doc.add_heading("Apollo.io Checklist", level=2)
apollo_items = [
    "Account created and verified",
    "Search filters configured (Industry, Title, SIC, NAICS, Location, Size, Revenue)",
    "Saved searches created for each target state",
    "Contact lists built and organized by state/priority",
    "Credits allocated and usage tracked",
    "Email verification enabled before export",
    "CRM integration connected (GHL or equivalent)",
    "Sequence templates loaded for cold outreach",
]
for item in apollo_items:
    add_bullet(doc, item)

# --- Google Maps / Yelp ---
doc.add_heading("Google Maps / Yelp Scraping Checklist", level=2)
gm_items = [
    "Chrome extensions installed (G Maps Extractor or equivalent)",
    "Outscraper account set up (if using)",
    "First city scraped as test run",
    "Data export tested (CSV format verified)",
    "Deduplication process established",
    "Data cleaning workflow set up (remove duplicates, verify emails)",
    "Lead import to GHL tested",
    "Scraping schedule created (2\u20133 cities per day)",
]
for item in gm_items:
    add_bullet(doc, item)

doc.add_page_break()

# =========================================================================
# SECTION 6: WEEKLY KPI TARGETS
# =========================================================================
doc.add_heading("SECTION 6: WEEKLY KPI TARGETS", level=1)

add_body(doc, "The following table outlines target KPIs across a 6-week launch campaign. Adjust based on actual response rates and conversion data.")

doc.add_paragraph()

kpi_headers = ["Metric", "Week 1\u20132", "Week 3\u20134", "Week 5", "Week 6"]
kpi_rows = [
    ("IG DMs sent/day", "30\u201350", "20\u201330", "15\u201320", "10 (re-engage)"),
    ("Cold emails sent/day (GHL)", "150\u2013200", "50 nurture", "50 last-chance", "50 close"),
    ("LinkedIn requests/day", "20\u201330", "15\u201320", "10\u201315", "5\u201310"),
    ("Facebook group posts/week", "0 (comment only)", "3\u20135", "3\u20135", "2\u20133"),
    ("Landing page visits", "50\u2013100", "200\u2013400", "400+", "400+ (retarget)"),
    ("Email open rate", "\u2014", "35\u201345%", "40\u201355%", "40\u201355%"),
    ("Conversions (sales)", "0\u20135 early", "10\u201320", "15\u201325", "15\u201330"),
]
add_styled_table(doc, kpi_headers, kpi_rows, col_widths=[2.0, 1.3, 1.3, 1.3, 1.3])

doc.add_paragraph()

add_sub_heading(doc, "Cumulative Targets")
add_bullet(doc, "50\u2013100 sales ($4,850 \u2013 $9,700 revenue at $97/sale)", bold_prefix="6-Week Goal: ")
add_bullet(doc, "5,000\u201310,000 total leads scraped/sourced", bold_prefix="Total Leads: ")
add_bullet(doc, "2,000\u20134,000 cold emails sent", bold_prefix="Total Emails: ")
add_bullet(doc, "1,500\u20132,500 DMs sent across IG + LinkedIn", bold_prefix="Total DMs: ")

doc.add_page_break()

# =========================================================================
# SECTION 7: PROFESSIONAL ASSOCIATIONS & INDUSTRY RESOURCES
# =========================================================================
doc.add_heading("SECTION 7: PROFESSIONAL ASSOCIATIONS & INDUSTRY RESOURCES", level=1)

add_sub_heading(doc, "Professional Associations")

assoc_headers = ["Organization", "Website", "Value"]
assoc_rows = [
    ("Professional Beauty Association (PBA)", "probeauty.org", "Largest US beauty trade association; events, education, advocacy"),
    ("International SalonSpa Business Network (ISBN)", "salonspanetwork.org", "Peer-to-peer network for salon/spa owners; benchmarking data"),
    ("Associated Hair Professionals (AHP)", "associatedhairprofessionals.com", "Membership org; insurance, education, community for stylists/owners"),
]
add_styled_table(doc, assoc_headers, assoc_rows, col_widths=[2.5, 2.2, 2.5])

doc.add_paragraph()

add_sub_heading(doc, "Industry Publications & Media")

pub_headers = ["Publication", "Website", "Focus"]
pub_rows = [
    ("Salon Today", "salontoday.com", "Business management, marketing, industry trends"),
    ("Modern Salon", "modernsalon.com", "Professional education, trends, salon business"),
    ("American Salon", "americansalonmag.com", "Industry news, techniques, salon management"),
]
add_styled_table(doc, pub_headers, pub_rows, col_widths=[2.0, 2.5, 2.8])

doc.add_paragraph()

add_sub_heading(doc, "Why These Matter for Lead Research")
add_bullet(doc, "PBA events and member directories can surface salon owner contacts.")
add_bullet(doc, "ISBN members are business-focused owners \u2014 ideal target profile.")
add_bullet(doc, "AHP members are engaged professionals investing in their careers/businesses.")
add_bullet(doc, "Industry publications feature salon owners \u2014 profile articles contain names, locations, and business details.")
add_bullet(doc, "Trade show attendee lists and speaker directories are additional lead sources.")

# =========================================================================
# SAVE
# =========================================================================
output_path = "/Users/nishamore/Documents/Salon world wide/Lead_Research_Bible.docx"
doc.save(output_path)
print(f"Document saved successfully to: {output_path}")
print(f"File size: {os.path.getsize(output_path):,} bytes")
