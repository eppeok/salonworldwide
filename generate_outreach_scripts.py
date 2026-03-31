#!/usr/bin/env python3
"""
Generate Complete Outreach Scripts DOCX for Salon Marketing Masterclass.
"""

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

OUTPUT_PATH = "/Users/nishamore/Documents/Salon world wide/Complete_Outreach_Scripts.docx"

doc = Document()

# ── Global defaults ──────────────────────────────────────────────────────
style = doc.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(11)
font.color.rgb = RGBColor(0x2D, 0x2D, 0x2D)
style.paragraph_format.space_after = Pt(6)
style.paragraph_format.line_spacing = 1.15

# ── Heading styles ───────────────────────────────────────────────────────
for level, (size, color, bold) in {
    1: (Pt(24), RGBColor(0x1A, 0x1A, 0x2E), True),
    2: (Pt(18), RGBColor(0x2C, 0x3E, 0x50), True),
    3: (Pt(14), RGBColor(0x8E, 0x44, 0xAD), True),
}.items():
    h = doc.styles[f'Heading {level}']
    h.font.name = 'Calibri'
    h.font.size = size
    h.font.color.rgb = color
    h.font.bold = bold
    h.paragraph_format.space_before = Pt(18 if level == 1 else 14 if level == 2 else 10)
    h.paragraph_format.space_after = Pt(8)

# Margins
for section in doc.sections:
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)


# ── Helper functions ─────────────────────────────────────────────────────

def add_page_break():
    doc.add_page_break()

def add_divider():
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("─" * 60)
    run.font.color.rgb = RGBColor(0xBD, 0xBD, 0xBD)
    run.font.size = Pt(10)

def add_label(text, color=RGBColor(0x2C, 0x3E, 0x50)):
    """Add a bold label paragraph."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = color
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    return p

def add_script_text(text):
    """Add script/message text in italic with a light background feel."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    run.italic = True
    run.font.size = Pt(10.5)
    run.font.color.rgb = RGBColor(0x34, 0x49, 0x5E)
    return p

def add_script_block(label_text, script_text, label_color=RGBColor(0x2C, 0x3E, 0x50)):
    """Add a labeled script block (bold label + italic script)."""
    add_label(label_text, label_color)
    add_script_text(script_text)

def add_email(subject_lines, preview, body_text, signature="[Your Name]"):
    """Add a formatted email block."""
    # Subject lines
    for subj in subject_lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        run_label = p.add_run("Subject: ")
        run_label.bold = True
        run_label.font.size = Pt(10.5)
        run_label.font.color.rgb = RGBColor(0x8E, 0x44, 0xAD)
        run_val = p.add_run(subj)
        run_val.font.size = Pt(10.5)
        run_val.font.color.rgb = RGBColor(0x2D, 0x2D, 0x2D)

    # Preview
    if preview:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        run_label = p.add_run("Preview: ")
        run_label.bold = True
        run_label.font.size = Pt(10.5)
        run_label.font.color.rgb = RGBColor(0x8E, 0x44, 0xAD)
        run_val = p.add_run(preview)
        run_val.italic = True
        run_val.font.size = Pt(10.5)
        run_val.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)

    add_divider()

    # Body
    for para_text in body_text.strip().split("\n\n"):
        lines = para_text.strip().split("\n")
        for line in lines:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.2)
            p.paragraph_format.space_after = Pt(4)
            # Check if line starts with a bullet marker
            if line.strip().startswith("- "):
                p.paragraph_format.left_indent = Inches(0.5)
                run = p.add_run("  " + line.strip())
            else:
                run = p.add_run(line.strip())
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0x34, 0x49, 0x5E)
            run.italic = True

    # Spacer
    doc.add_paragraph()


def add_shaded_box(text):
    """Add text with shading to simulate a box."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.right_indent = Inches(0.3)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    # Add shading
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F4F6F7"/>')
    p.paragraph_format.element.get_or_add_pPr().append(shading_elm)
    run = p.add_run(text)
    run.font.size = Pt(10.5)
    run.font.color.rgb = RGBColor(0x2D, 0x2D, 0x2D)
    return p


# ═════════════════════════════════════════════════════════════════════════
# COVER PAGE
# ═════════════════════════════════════════════════════════════════════════

for _ in range(6):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("COMPLETE OUTREACH SCRIPTS")
run.bold = True
run.font.size = Pt(36)
run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)
run.font.name = 'Calibri'

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(12)
run = p.add_run("Every Message, Every Channel, Every Phase")
run.font.size = Pt(18)
run.font.color.rgb = RGBColor(0x8E, 0x44, 0xAD)
run.italic = True

add_divider()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(16)
run = p.add_run("Salon Marketing Masterclass")
run.font.size = Pt(14)
run.font.color.rgb = RGBColor(0x2C, 0x3E, 0x50)
run.bold = True

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("evolvxai.com/salon01-3329")
run.font.size = Pt(13)
run.font.color.rgb = RGBColor(0x8E, 0x44, 0xAD)

for _ in range(4):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("$97 Launch Price  |  8 Video Modules  |  5+ Hours  |  30-Day Money-Back Guarantee")
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("500+ Salon Owners Enrolled  |  6 Bonus Toolkits Worth $1,041")
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)


# ═════════════════════════════════════════════════════════════════════════
# PART 1: INSTAGRAM DM SCRIPTS
# ═════════════════════════════════════════════════════════════════════════
add_page_break()
doc.add_heading("PART 1: INSTAGRAM DM SCRIPTS", level=1)

# Phase 1 — Cold Openers
doc.add_heading("Phase 1 — Cold Openers (Week 1-2)", level=2)

add_script_block(
    "Script A — The Ad Question:",
    "Hey [Name] — noticed you're running ads for your salon. Quick question: are they actually converting? I work with salon owners on this exact problem and I'm curious what your experience has been."
)

add_script_block(
    "Script B — The Compliment:",
    "Hey [Name] — your salon looks amazing! Quick question — do most of your new clients find you through Instagram or ads? I've been researching how US salon owners are getting clients and your page caught my eye."
)

add_script_block(
    "Script C — The Curiosity:",
    "Hey [Name] — random question: if you could fill your calendar without spending on ads, would that change your business? I've been talking to a lot of salon owners in [their state] about this lately."
)

add_script_block(
    "Script D — The Value-First:",
    "Hey [Name] — I just shared a free tip with a salon owner in [their city] about ranking #1 on Google Maps for 'salon near me' — took her 2 weeks. Want me to share it with you too?"
)

# Phase 1 — Follow-Up Replies
doc.add_heading("Phase 1 — Follow-Up Replies (When They Respond)", level=2)

add_script_block(
    "Reply A — If ads aren't working:",
    "Yeah that's the trap most salon owners fall into — paying to rent clients instead of building a system. I actually just wrapped up a masterclass on this — over 500 salon owners have used it to stop running ads entirely. Happy to send you the link if you want to check it out."
)

add_script_block(
    "Reply B — If they ask what you do:",
    "I help salon owners fill their calendars using organic marketing — Google Maps, Instagram, client reactivation, and referrals. No paid ads. I put everything into an 8-module masterclass that 500+ salon owners have used. It's $97 right now (normally $497) with a 30-day guarantee. Want me to send the link?"
)

add_script_block(
    "Reply C — If ads ARE working:",
    "Nice! That's rare honestly — most salon owners I talk to are bleeding $3k-$8k/month with inconsistent results. Out of curiosity, what would you do with that budget if you didn't need it for ads? I've been helping salon owners build organic systems that replace ads entirely. Might be worth a look even as a backup plan."
)

add_script_block(
    "Reply D — If lukewarm/maybe:",
    "No pressure at all. If you want, I can send you a quick breakdown of the 4-pillar system — takes 2 minutes to read. Then you can decide if it's worth exploring. Sound fair?"
)

# Phase 2 — Nurture DMs
doc.add_heading("Phase 2 — Nurture DMs (Week 3-4)", level=2)

add_script_block(
    "Story Reply Trigger:",
    "Hey [Name] — saw you watched my story about [topic]. Did that resonate with you? Happy to chat about it if you want."
)

add_script_block(
    "Value-Add DM:",
    "Hey [Name] — just wanted to share something quick. One of the salon owners I work with ranked #1 on Google Maps in [their city] in just 30 days. The trick? She optimized 3 things most salon owners ignore. Happy to share if you're curious."
)

# Phase 3 — Urgency DMs
doc.add_heading("Phase 3 — Urgency DMs (Week 5)", level=2)

add_script_block(
    "Urgency Message:",
    "Hey [Name] — just wanted to give you a heads up. The $97 launch price on the Salon Marketing Masterclass closes [date]. After that it goes back to $497. If you've been thinking about it, now's the time. The 30-day guarantee means there's zero risk. Here's the link: evolvxai.com/salon01-3329"
)

# Phase 4 — Last Chance DMs
doc.add_heading("Phase 4 — Last Chance DMs (Week 6)", level=2)

add_script_block(
    "Closer:",
    "Hey [Name] — the 60% launch price on my salon marketing course closes [date]. Wanted to give you a personal heads up before I pull it. The 30-day money-back guarantee means there's zero risk. Just wanted to make sure you saw it before it was gone."
)

add_script_block(
    "Re-engage Non-Responders:",
    "Hey [Name] — I reached out a few weeks ago about salon marketing. Just wanted to check in one more time — I've had 500+ salon owners go through the masterclass and the results have been incredible. The launch price ($97) closes [date]. After that it's $497. No pressure, just didn't want you to miss it."
)


# ═════════════════════════════════════════════════════════════════════════
# PART 2: LINKEDIN OUTREACH SCRIPTS
# ═════════════════════════════════════════════════════════════════════════
add_page_break()
doc.add_heading("PART 2: LINKEDIN OUTREACH SCRIPTS", level=1)

doc.add_heading("Connection Request Notes (Under 300 Characters)", level=2)

add_script_block(
    "Note A:",
    "Hi [Name] — I help salon owners fill their calendars without paid ads. Would love to connect and share some insights that have worked for 500+ US salon owners."
)

add_script_block(
    "Note B:",
    "Hi [Name] — noticed you run a salon in [City]. I've been working with salon owners on organic marketing strategies. Would love to connect!"
)

add_script_block(
    "Note C:",
    "Hi [Name] — I just helped a salon owner in [State] rank #1 on Google Maps in 30 days. Happy to share how if you're interested. Let's connect!"
)

doc.add_heading("First Message After Connection (Day 1-2)", level=2)

add_script_text(
    "Hey [Name], thanks for connecting! Quick question — are you currently using paid ads to get clients for your salon? I've been working with salon owners across the US on a completely different approach and I'm curious where you stand."
)

doc.add_heading("Follow-Up (Day 3-5 If They Reply Interested)", level=2)

add_script_text(
    "That's what I hear from most salon owners. The system I teach has 4 pillars — Google Maps domination, Instagram content that converts, client reactivation, and referrals. 500+ salon owners have gone through it. It's all inside an 8-module masterclass — $97 right now (normally $497) with a 30-day guarantee. Want me to send the link?"
)

doc.add_heading("InMail Template", level=2)

add_label("Subject:", RGBColor(0x8E, 0x44, 0xAD))
add_script_text(
    "Salon owners in [City] are ditching paid ads — here's what they're doing instead"
)

add_label("Body:", RGBColor(0x8E, 0x44, 0xAD))
add_script_text(
    "Hi [Name], I've been working with salon owners across the US who were spending $3k-$8k/month on ads with inconsistent results. The ones who switched to organic marketing are now getting more bookings than ever — at $0 ad spend. I put everything into an 8-module masterclass (500+ salon owners enrolled). It's $97 during launch (normally $497) with a 30-day money-back guarantee. Would love to share it with you: evolvxai.com/salon01-3329"
)


# ═════════════════════════════════════════════════════════════════════════
# PART 3: FACEBOOK GROUP POST SCRIPTS
# ═════════════════════════════════════════════════════════════════════════
add_page_break()
doc.add_heading("PART 3: FACEBOOK GROUP POST SCRIPTS", level=1)

doc.add_heading("Week 1 — Comment Scripts (NO Posts, Only Comments)", level=2)

add_script_block(
    "Comment if someone asks about marketing:",
    "Honestly, the biggest ROI I've seen salon owners get is from Google Maps optimization + client reactivation campaigns. Most salon owners overlook these because they're not 'sexy' like Instagram ads, but they convert way better."
)

add_script_block(
    "Comment if someone complains about ads:",
    "This is so common. I've talked to hundreds of salon owners and the $3k-$8k/month ad spend with inconsistent results is almost universal. The ones who broke free built organic systems instead."
)

doc.add_heading("Week 2+ — Organic Posts (NO Links in Post Body)", level=2)

add_script_block(
    "Post 1 — The Ad Trap:",
    "Honest question for anyone running paid ads: how much are you spending per month? I've been talking to salon owners across the US and the number is almost always between $3k-$8k. Most say the results are inconsistent. The ones who got off the ad treadmill all have one thing in common — they built a 4-pillar organic system instead. Happy to share what I've seen working if anyone wants to chat."
)

add_script_block(
    "Post 2 — The Reactivation Story:",
    "Quick story: A salon owner I know sent a 3-message WhatsApp sequence to clients who hadn't visited in 90+ days. 23 of them came back within a week. Zero ad spend. Zero awkwardness. Just the right words at the right time. If anyone wants to know what the messages said, drop a comment or DM me."
)

add_script_block(
    "Post 3 — Google Maps Win:",
    "Random tip: One salon owner I work with went from page 3 to #1 on Google Maps for 'salon near me' in her city — in about 30 days. The 3 things she changed are things most salon owners completely ignore. If anyone wants the breakdown, let me know in the comments."
)

add_script_block(
    "Post 4 — The $48K Question:",
    "Did a quick calculation the other day. If you're spending $3k/month on ads, that's $36k/year. Add in the booking software fees, the social media manager, and the random 'boost this post' spending — you're looking at $48k+ per year just to stay visible. What if you could get MORE bookings for $0/month in ad spend? That's what I've been helping salon owners do."
)

add_script_block(
    "Post 5 — Referral System:",
    "Question for salon owners: How many referrals do you get per month? Not just 'word of mouth' but actual referrals from a system? Most salon owners I talk to say 'a few' but don't have a real system. The ones who do are getting 5-10 per month consistently. It's one of the 4 pillars I teach. Happy to share the basics if anyone's interested."
)

doc.add_heading("Comment CTA (When People Respond)", level=2)

add_script_text("DM me and I'll share the details!")


# ═════════════════════════════════════════════════════════════════════════
# PART 4: COLD EMAIL SEQUENCES FOR GHL
# ═════════════════════════════════════════════════════════════════════════
add_page_break()
doc.add_heading("PART 4: COLD EMAIL SEQUENCES FOR GHL", level=1)

p = doc.add_paragraph()
run = p.add_run("9-Email Sequence  |  GoHighLevel Personalization: {{contact.first_name}}")
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)
run.italic = True

add_divider()

# ─── EMAIL 1 ─────────────────────────────────────────────────────────────
doc.add_heading("Email 1 — The Pain (Phase 2, Day 1)", level=2)

add_email(
    subject_lines=[
        '"Are you renting clients or building a business?"',
        '"The $48,000 mistake most salon owners make every year"',
    ],
    preview='"Most salon owners don\'t realize the trap they\'re in until they see this number..."',
    body_text="""Hey {{contact.first_name}},

Quick math for you.

If you're spending $3,000/month on ads for your salon — which is on the low end for most US salons — that's $36,000 a year.

Now ask yourself this: if you stopped paying tomorrow, would your bookings dry up?

If the answer is yes, you're not building a business. You're renting clients.

I've talked to hundreds of salon owners who've been stuck on this treadmill for years. The worst part? The moment the ads stop, everything stops. There's no system underneath. No loyalty. No referrals. Just silence.

The good news: there is a way out. And it doesn't require more ad spend. It actually requires zero ad spend.

I'll show you what it looks like in my next email.

Talk soon,
[Your Name]

P.S. — The number most salon owners discover when they finally do the math? $48,000+/year gone. Just in ads. With inconsistent results."""
)

add_divider()

# ─── EMAIL 2 ─────────────────────────────────────────────────────────────
doc.add_heading("Email 2 — Social Proof (Phase 2, Day 3)", level=2)

add_email(
    subject_lines=[
        '"23 clients came back. Zero ad spend."',
        '"How Fatima stopped paying $4,000/month for ads (and got MORE bookings)"',
    ],
    preview='"Maria thought those clients were gone forever. Then she ran one simple campaign..."',
    body_text="""Hey {{contact.first_name}},

Yesterday I talked about the ad treadmill — spending $3k-$8k/month just to stay visible.

Today I want to show you what happens when salon owners get off it.

Meet Fatima Al-Rashid:
She was spending $4,000/month on Instagram ads. Barely any return. After implementing just 3 tactics from a system I'll share with you, she got more bookings than ever — and completely stopped paid ads.
Current ad spend: $0. Extra bookings per month: +35.

Meet Maria Petrova:
The reactivation campaign brought back 23 clients she thought were gone forever. Using simple, non-pushy WhatsApp scripts. Week 1. The course paid for itself immediately.

These aren't outliers. This is what happens when you build a real marketing system instead of renting traffic.

The system has 4 pillars:
- 1. RANK — dominate Google Maps so clients find YOU first
- 2. CONVERT — Instagram content that turns followers into bookings
- 3. REACTIVATE — bring back 30-45% of clients who haven't visited in 90+ days
- 4. REFER — simple system that generates 5-10 referrals per month

Tomorrow I'll show you exactly what's inside.

[Your Name]"""
)

add_divider()

# ─── EMAIL 3 ─────────────────────────────────────────────────────────────
doc.add_heading("Email 3 — The Tease (Phase 2, Day 6)", level=2)

add_email(
    subject_lines=[
        '"The 8-module system that filled 500+ salon calendars (sneak peek)"',
        '"Before I show you the price — here\'s what\'s inside"',
    ],
    preview='"Over the last two days I introduced you to the problem and the proof. Today, the system."',
    body_text="""Hey {{contact.first_name}},

Over the last two emails I introduced you to the problem (the ad treadmill) and the proof (Fatima, Maria, and 500+ others).

Today I want to give you a peek at the actual system.

The Salon Marketing Masterclass is 8 video modules — 5+ hours of step-by-step training specifically built for salon owners. No fluff. No theory. Just tactics you can implement the same week.

Here's what's inside:
- Module 1 — Foundation & Quick Wins Audit
- Module 2 — Google Business Domination (rank #1 for 'salon near me')
- Module 3 — Instagram Content That Converts (30-day calendar included)
- Module 4 — Before/After Content System
- Module 5 — Client Reactivation Campaign (the 3-wave system)
- Module 6 — Referral & Partnership Strategy
- Module 7 — Promotions & QR Capture
- Module 8 — Review Amplification (turn every 5-star review into 7 assets)

Plus 6 implementation toolkits — templates, scripts, and checklists worth $544 on their own.

In my next email, I'll send you the full offer and the price.

Spoiler: it's not $497.

[Your Name]"""
)

add_divider()

# ─── EMAIL 4 ─────────────────────────────────────────────────────────────
doc.add_heading("Email 4 — Offer Reveal (Phase 3, Day 1)", level=2)

add_email(
    subject_lines=[
        '"Here\'s everything you get — and the price drops today"',
        '"$1,041 in value. $97 today. Here\'s the full breakdown."',
    ],
    preview='"The full Salon Marketing Masterclass is open. Here\'s what\'s included..."',
    body_text="""Hey {{contact.first_name}},

I promised you the full offer today. Here it is.

The Salon Marketing Masterclass includes:
- 8 Video Modules (5+ hours of step-by-step training) — $497 value
- 30-Day Action Planner — $89 value
- GBP Domination Kit (100+ photo ideas, templates, schedule) — $79 value
- Instagram Content Vault (90 days of captions + Canva templates) — $99 value
- Reactivation Swipe File (copy-paste WhatsApp + SMS scripts) — $99 value
- 12-Month Seasonal Calendar — $79 value
- Review Amplification Pack — $99 value

Total value: $1,041.

Launch price today: $97.

One-time payment. Lifetime access. 30-day money-back guarantee — implement the tactics, don't see results, email us with proof and you get every dollar back.

Enroll now: evolvxai.com/salon01-3329

This price does not last. I'll tell you exactly when it ends in my next email.

[Your Name]

P.S. — 500+ salon owners have already gone through this. The only common thing among the ones who didn't see results? They never implemented."""
)

add_divider()

# ─── EMAIL 5 ─────────────────────────────────────────────────────────────
doc.add_heading("Email 5 — Bonus Stack (Phase 3, Day 3)", level=2)

add_email(
    subject_lines=[
        '"Here\'s what you lose if you wait (the bonuses expire too)"',
        '"The bonuses disappear before the price goes up"',
    ],
    preview='"It\'s not just the price that has a deadline..."',
    body_text="""Hey {{contact.first_name}},

Yesterday I showed you the full offer — $1,041 in training and toolkits for $97.

Today I want to make sure you know: it's not just the price that has a deadline.

The launch bonuses — the Reactivation Swipe File, the Instagram Content Vault, the 30-Day Action Planner — these are included at $97 only for the launch window. After that, the course is still available, but the bonuses stack goes away.

Why does this matter?

Because the Reactivation Swipe File alone is the asset that brought Maria 23 clients back in Week 1. That's not a nice-to-have. That's direct revenue the day you implement it.

The Instagram Content Vault gives you 90 days of captions and Canva templates — so you never stare at a blank post again.

The 30-Day Action Planner tells you exactly what to do every morning. No guessing.

All of this disappears when the launch window closes.

Enroll before that happens: evolvxai.com/salon01-3329

[Your Name]"""
)

add_divider()

# ─── EMAIL 6 ─────────────────────────────────────────────────────────────
doc.add_heading("Email 6 — Testimonials (Phase 3, Day 6)", level=2)

add_email(
    subject_lines=[
        '"Real salon owners. Real numbers. Read these."',
        '"From $4,000/month in ads to $0 — their stories"',
    ],
    preview='"Before you decide, hear from the people who\'ve already done this..."',
    body_text="""Hey {{contact.first_name}},

I'm going to keep this one short and let the results do the talking.

"I was spending $4,000/month on Instagram ads with barely any return. After implementing just 3 tactics from this course, I'm getting more bookings than ever and I've completely stopped paid ads."
— Fatima Al-Rashid | Ad spend now: $0 | Extra bookings: +35/month

"The reactivation campaign brought back 23 clients I thought were gone forever. The WhatsApp scripts are gold. Simple, not pushy, and they actually work."
— Maria Petrova | 23 clients reactivated | Paid for itself in Week 1

500+ salon owners have enrolled. These aren't the exceptions — they're what happens when you actually implement.

You're 30 days away from results like these. The only thing standing between you and a full calendar is the decision.

Enroll here: evolvxai.com/salon01-3329

[Your Name]

P.S. — The launch price closes soon. Don't be the person who checks this email after it's gone and wishes they'd moved faster."""
)

add_divider()

# ─── EMAIL 7 ─────────────────────────────────────────────────────────────
doc.add_heading("Email 7 — 48hr Warning (Phase 4, Day 1)", level=2)

add_email(
    subject_lines=[
        '"48 hours left — then the price goes back to $497"',
        '"The math on waiting vs. enrolling right now"',
    ],
    preview='"The launch window is closing. Here\'s what you need to know..."',
    body_text="""Hey {{contact.first_name}},

Quick heads up: the $97 launch price for Salon Marketing Masterclass closes in 48 hours.

After that, it goes back to $497.

I'm not saying this to pressure you. I'm saying it because I've seen what happens when salon owners wait — they end up spending another $4,000-$8,000 on ads next month, and then the month after that, while this system sits one click away.

What you get for $97 today:
- 8 modules — 5+ hours of implementation-ready training
- 6 bonus toolkits — templates, scripts, and calendars
- 30-day money-back guarantee — zero risk
- Lifetime access with free updates

Enroll now before the window closes: evolvxai.com/salon01-3329

[Your Name]"""
)

add_divider()

# ─── EMAIL 8 ─────────────────────────────────────────────────────────────
doc.add_heading("Email 8 — 24hr Warning (Phase 4, Day 2)", level=2)

add_email(
    subject_lines=[
        '"Last chance — 24 hours before the price changes"',
        '"What if it doesn\'t work? (Read this guarantee)"',
    ],
    preview='"This is the final reminder. After midnight tomorrow, the $97 price is gone..."',
    body_text="""Hey {{contact.first_name}},

24 hours.

That's how long the $97 price is still available for Salon Marketing Masterclass.

Tomorrow it goes back to $497.

I want to answer the most common objection I hear at this point:

"What if it doesn't work for my salon?"

That's exactly why there's a 30-day money-back guarantee. Implement the tactics for 30 days. If you don't see a measurable increase in bookings, email us with proof you did the work — you get a full refund. No questions asked.

There is genuinely no risk here. The only risk is spending another month paying for ads while the system that replaces them costs $97.

Last chance: evolvxai.com/salon01-3329

[Your Name]

P.S. — 500+ salon owners already made this decision. You know what they all have in common now? Full calendars."""
)

add_divider()

# ─── EMAIL 9 ─────────────────────────────────────────────────────────────
doc.add_heading("Email 9 — Final Close (Phase 4, Day 3 — 6hrs Before)", level=2)

add_email(
    subject_lines=[
        '"This is it — closing in 6 hours [FINAL NOTICE]"',
        '"Final email. Then it\'s done."',
    ],
    preview='"I\'m closing the cart in 6 hours. This is the last email I\'ll send about this..."',
    body_text="""Hey {{contact.first_name}},

I'm closing the $97 launch price in 6 hours.

This is the last email I'll send about this. After tonight, the price goes to $497 and the launch bonuses are gone.

If you've been on the fence — this is the moment.

I don't believe in manufactured urgency. This deadline is real. The system works. The guarantee removes all risk.

If you want to stop bleeding money on ads and build a salon calendar that fills itself organically — this is how.

Final link: evolvxai.com/salon01-3329

If you're not ready, no hard feelings. But if you are — move now.

[Your Name]

P.S. — After the cart closes tonight, I'll be back with new content and new topics. But this offer — $97, full bonuses, this guarantee — won't come back at this price. This is it."""
)


# ═════════════════════════════════════════════════════════════════════════
# PART 5: WHATSAPP FOLLOW-UP SCRIPTS
# ═════════════════════════════════════════════════════════════════════════
add_page_break()
doc.add_heading("PART 5: WHATSAPP FOLLOW-UP SCRIPTS", level=1)

doc.add_heading("Voice Note Scripts", level=2)

add_script_block(
    "Voice Note 1 — Warm Lead Follow-Up:",
    "Hey [Name], just wanted to follow up personally — I know we spoke about your salon marketing a few weeks back. The Salon Marketing Masterclass is closing the launch price on [date] and I wanted to give you a heads up before I pull it. It's $97 right now, goes back to $497 after that. Happy to answer any questions before you decide — just reply here."
)

add_script_block(
    "Voice Note 2 — Post-Chat Follow-Up:",
    "Hey [Name], it's [Your Name]. I wanted to check in after our chat. Did the 4-pillar system make sense for your salon? The launch price closes [date] — just wanted to make sure you had all the info you needed before deciding. No pressure at all."
)

doc.add_heading("Text Message Scripts", level=2)

add_script_block(
    "Text Message 1 — Link Drop:",
    "Hey [Name]! Here's the link to the Salon Marketing Masterclass we talked about: evolvxai.com/salon01-3329 — $97 (normally $497), 30-day guarantee. Let me know if you have any questions!"
)

add_script_block(
    "Text Message 2 — Final Reminder:",
    "Hey [Name] — quick heads up: the $97 price closes tonight. After that it's $497. Just didn't want you to miss it. Link: evolvxai.com/salon01-3329"
)


# ═════════════════════════════════════════════════════════════════════════
# PART 6: THREADS OUTREACH SCRIPTS
# ═════════════════════════════════════════════════════════════════════════
add_page_break()
doc.add_heading("PART 6: THREADS OUTREACH SCRIPTS", level=1)

doc.add_heading("Reply Scripts (Engaging on Salon Owner Posts)", level=2)

add_script_text(
    "This is so true — I've been working with salon owners on exactly this. The ones who switched to organic marketing are spending $0 on ads and getting more bookings than before."
)

doc.add_heading("Original Post Scripts", level=2)

add_script_text(
    "Hot take: salon owners spending $3k-$8k/month on ads are renting their clients, not building a business. The smart ones are building organic systems. I just helped 500+ do exactly that."
)


# ═════════════════════════════════════════════════════════════════════════
# SAVE
# ═════════════════════════════════════════════════════════════════════════
doc.save(OUTPUT_PATH)
print(f"Document saved successfully to: {OUTPUT_PATH}")
print(f"File size: {os.path.getsize(OUTPUT_PATH):,} bytes")
