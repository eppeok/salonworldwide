from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import os

doc = Document()
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)

def add_heading(text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x1B, 0x3A, 0x5C)
    return h

def add_script_block(label, text, is_italic=True):
    """Add a labeled script block with visual distinction"""
    p = doc.add_paragraph()
    run = p.add_run(label)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
    
    p2 = doc.add_paragraph()
    run2 = p2.add_run(text)
    if is_italic:
        run2.italic = True
    run2.font.size = Pt(10)
    # Add left indent for visual distinction
    p2.paragraph_format.left_indent = Inches(0.3)
    return p2

def add_email_block(num, name, phase, subj_a, subj_b, preview, body, ps=""):
    add_heading(f"Email {num} — {name} ({phase})", 3)
    p = doc.add_paragraph()
    r = p.add_run(f"Subject Line A: ")
    r.bold = True
    r.font.size = Pt(10)
    r2 = p.add_run(subj_a)
    r2.font.size = Pt(10)
    
    p = doc.add_paragraph()
    r = p.add_run(f"Subject Line B: ")
    r.bold = True
    r.font.size = Pt(10)
    r2 = p.add_run(subj_b)
    r2.font.size = Pt(10)
    
    p = doc.add_paragraph()
    r = p.add_run(f"Preview Text: ")
    r.bold = True
    r.font.size = Pt(10)
    r2 = p.add_run(preview)
    r2.italic = True
    r2.font.size = Pt(10)
    
    doc.add_paragraph()  # spacer
    for line in body.strip().split('\n'):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.3)
        r = p.add_run(line)
        r.font.size = Pt(10)
    
    if ps:
        doc.add_paragraph()
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.3)
        r = p.add_run(ps)
        r.italic = True
        r.font.size = Pt(10)

# ===================== COVER PAGE =====================
for _ in range(6):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("COMPLETE OUTREACH SCRIPTS")
run.bold = True
run.font.size = Pt(36)
run.font.color.rgb = RGBColor(0x1B, 0x3A, 0x5C)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Every Message, Every Channel, Every Phase")
run.font.size = Pt(16)
run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Salon Marketing Masterclass  |  evolvxai.com/salon01-3329")
run.font.size = Pt(12)
run.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("GHL Merge Tag: {{contact.first_name}}")
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

doc.add_page_break()

# ===================== PART 1: INSTAGRAM DM SCRIPTS =====================
add_heading("PART 1: INSTAGRAM DM SCRIPTS", 1)

add_heading("Phase 1 — Cold Openers (Week 1-2)", 2)
doc.add_paragraph("Send 30-50 of these per day. Mix and match variants. NEVER pitch in the first message. These are conversation starters only.")

add_script_block(
    "Script A — The Ad Question Opener:",
    '"Hey [Name] \u2014 noticed you\'re running ads for your salon. Quick question: are they actually converting? I work with salon owners on this exact problem and I\'m curious what your experience has been."'
)

add_script_block(
    "Script B — The Compliment Opener:",
    '"Hey [Name] \u2014 your salon looks amazing! Quick question \u2014 do most of your new clients find you through Instagram or ads? I\'ve been researching how US salon owners are getting clients and your page caught my eye."'
)

add_script_block(
    "Script C — The Curiosity Opener:",
    '"Hey [Name] \u2014 random question: if you could fill your calendar without spending on ads, would that change your business? I\'ve been talking to a lot of salon owners in [their state] about this lately."'
)

add_script_block(
    "Script D — The Value-First Opener:",
    '"Hey [Name] \u2014 I just shared a free tip with a salon owner in [their city] about ranking #1 on Google Maps for \'salon near me\' \u2014 took her 2 weeks. Want me to share it with you too?"'
)

add_heading("Phase 1 — Follow-Up Replies (When They Respond)", 2)

add_script_block(
    "Reply A — If ads AREN'T working:",
    '"Yeah that\'s the trap most salon owners fall into \u2014 paying to rent clients instead of building a system. I actually just wrapped up a masterclass on this \u2014 over 500 salon owners have used it to stop running ads entirely. Happy to send you the link if you want to check it out."'
)

add_script_block(
    "Reply B — If they ask what you do:",
    '"I help salon owners fill their calendars using organic marketing \u2014 Google Maps, Instagram, client reactivation, and referrals. No paid ads. I put everything into an 8-module masterclass that 500+ salon owners have used. It\'s $97 right now (normally $497) with a 30-day guarantee. Want me to send the link?"'
)

add_script_block(
    "Reply C — If ads ARE working:",
    '"Nice! That\'s rare honestly \u2014 most salon owners I talk to are bleeding $3k-$8k/month with inconsistent results. Out of curiosity, what would you do with that budget if you didn\'t need it for ads? I\'ve been helping salon owners build organic systems that replace ads entirely. Might be worth a look even as a backup plan."'
)

add_script_block(
    "Reply D — If lukewarm / maybe:",
    '"No pressure at all. If you want, I can send you a quick breakdown of the 4-pillar system \u2014 takes 2 minutes to read. Then you can decide if it\'s worth exploring. Sound fair?"'
)

add_script_block(
    "Reply E — If they say YES send the link:",
    '"Here you go: evolvxai.com/salon01-3329 \u2014 it\'s $97 right now (normally $497) and comes with a 30-day money-back guarantee. The Reactivation Swipe File alone has helped salon owners bring back 20+ clients in Week 1. Let me know if you have any questions!"'
)

add_heading("Phase 2 — Nurture DMs (Week 3-4)", 2)

add_script_block(
    "Story Reply Trigger (when they view your stories):",
    '"Hey [Name] \u2014 saw you watched my story about [topic]. Did that resonate with you? Happy to chat about it if you want."'
)

add_script_block(
    "Value-Add DM:",
    '"Hey [Name] \u2014 just wanted to share something quick. One of the salon owners I work with ranked #1 on Google Maps in [their city] in just 30 days. The trick? She optimized 3 things most salon owners ignore. Happy to share if you\'re curious."'
)

add_script_block(
    "Check-In DM:",
    '"Hey [Name] \u2014 just checking in! Have you made any changes to your marketing since we last chatted? I\'ve been sharing a lot about the 4-pillar system on my stories \u2014 would love to hear if anything clicked for you."'
)

add_heading("Phase 3 — Urgency DMs (Week 5)", 2)

add_script_block(
    "Urgency DM:",
    '"Hey [Name] \u2014 just wanted to give you a heads up. The $97 launch price on the Salon Marketing Masterclass closes [date]. After that it goes back to $497. If you\'ve been thinking about it, now\'s the time. The 30-day guarantee means there\'s zero risk. Here\'s the link: evolvxai.com/salon01-3329"'
)

add_heading("Phase 4 — Last Chance DMs (Week 6)", 2)

add_script_block(
    "Closer DM:",
    '"Hey [Name] \u2014 the 60% launch price on my salon marketing course closes [date]. Wanted to give you a personal heads up before I pull it. The 30-day money-back guarantee means there\'s zero risk. Just wanted to make sure you saw it before it was gone."'
)

add_script_block(
    "Re-Engage Non-Responders:",
    '"Hey [Name] \u2014 I reached out a few weeks ago about salon marketing. Just wanted to check in one more time \u2014 I\'ve had 500+ salon owners go through the masterclass and the results have been incredible. The launch price ($97) closes [date]. After that it\'s $497. No pressure, just didn\'t want you to miss it."'
)

doc.add_page_break()

# ===================== PART 2: LINKEDIN SCRIPTS =====================
add_heading("PART 2: LINKEDIN OUTREACH SCRIPTS", 1)

add_heading("Connection Request Notes (Under 300 Characters)", 2)

add_script_block(
    "Note A — Direct:",
    '"Hi [Name] \u2014 I help salon owners fill their calendars without paid ads. Would love to connect and share some insights that have worked for 500+ US salon owners."'
)

add_script_block(
    "Note B — Curiosity:",
    '"Hi [Name] \u2014 noticed you run a salon in [City]. I\'ve been working with salon owners on organic marketing strategies. Would love to connect!"'
)

add_script_block(
    "Note C — Value-First:",
    '"Hi [Name] \u2014 I just helped a salon owner in [State] rank #1 on Google Maps in 30 days. Happy to share how if you\'re interested. Let\'s connect!"'
)

add_heading("First Message After Connection (Day 1-2)", 2)

add_script_block(
    "Opening Message:",
    '"Hey [Name], thanks for connecting! Quick question \u2014 are you currently using paid ads to get clients for your salon? I\'ve been working with salon owners across the US on a completely different approach and I\'m curious where you stand."'
)

add_heading("Follow-Up Messages", 2)

add_script_block(
    "Follow-Up If Interested (Day 3-5):",
    '"That\'s what I hear from most salon owners. The system I teach has 4 pillars \u2014 Google Maps domination, Instagram content that converts, client reactivation, and referrals. 500+ salon owners have gone through it. It\'s all inside an 8-module masterclass \u2014 $97 right now (normally $497) with a 30-day guarantee. Want me to send the link?"'
)

add_script_block(
    "Follow-Up If No Reply (Day 7):",
    '"Hey [Name] \u2014 just wanted to follow up on my last message. No worries if you\'re busy! I\'ve been sharing some insights about salon marketing on my profile \u2014 feel free to check out my recent posts. Always happy to chat if you have questions."'
)

add_heading("LinkedIn InMail Template (Sales Navigator)", 2)

add_script_block(
    "Subject: \"Salon owners in [City] are ditching paid ads \u2014 here's what they're doing instead\"",
    '"Hi [Name], I\'ve been working with salon owners across the US who were spending $3k-$8k/month on ads with inconsistent results. The ones who switched to organic marketing are now getting more bookings than ever \u2014 at $0 ad spend.\n\nI put everything into an 8-module masterclass (500+ salon owners enrolled). It\'s $97 during launch (normally $497) with a 30-day money-back guarantee.\n\nWould love to share it with you: evolvxai.com/salon01-3329\n\nBest,\n[Your Name]"'
)

doc.add_page_break()

# ===================== PART 3: FACEBOOK GROUP POSTS =====================
add_heading("PART 3: FACEBOOK GROUP POST SCRIPTS", 1)

add_heading("Week 1 — Comment Scripts Only (NO Posts)", 2)
doc.add_paragraph("During Week 1, you ONLY comment on others' posts. Give genuine value. Build credibility. Zero promotion.")

add_script_block(
    "Comment — If someone asks about marketing:",
    '"Honestly, the biggest ROI I\'ve seen salon owners get is from Google Maps optimization + client reactivation campaigns. Most salon owners overlook these because they\'re not \'sexy\' like Instagram ads, but they convert way better."'
)

add_script_block(
    "Comment — If someone complains about ads:",
    '"This is so common. I\'ve talked to hundreds of salon owners and the $3k-$8k/month ad spend with inconsistent results is almost universal. The ones who broke free built organic systems instead."'
)

add_script_block(
    "Comment — If someone asks about getting more clients:",
    '"Three things I\'ve seen work consistently: (1) optimize your Google Business Profile so you rank for \'salon near me\' in your city, (2) reactivate clients who haven\'t visited in 90+ days \u2014 most salon owners are sitting on a goldmine of lapsed clients, and (3) build a simple referral system. None of this requires ads."'
)

add_heading("Week 2+ — Organic Posts (NO Links in Post Body)", 2)
doc.add_paragraph("Post 1 per group per week. Rotate across groups on Monday/Wednesday/Friday. NEVER include links in the post body. CTA is always 'DM me' or 'comment below.'")

add_script_block(
    "Post 1 — The Ad Trap:",
    '"Honest question for anyone running paid ads: how much are you spending per month?\n\nI\'ve been talking to salon owners across the US and the number is almost always between $3k\u2013$8k. Most say the results are inconsistent.\n\nThe ones who got off the ad treadmill all have one thing in common \u2014 they built a 4-pillar organic system instead.\n\nHappy to share what I\'ve seen working if anyone wants to chat."'
)

add_script_block(
    "Post 2 — The Reactivation Story:",
    '"Quick story: A salon owner I know sent a 3-message WhatsApp sequence to clients who hadn\'t visited in 90+ days.\n\n23 of them came back within a week.\n\nZero ad spend. Zero awkwardness. Just the right words at the right time.\n\nIf anyone wants to know what the messages said, drop a comment or DM me."'
)

add_script_block(
    "Post 3 — Google Maps Win:",
    '"Random tip: One salon owner I work with went from page 3 to #1 on Google Maps for \'salon near me\' in her city \u2014 in about 30 days.\n\nThe 3 things she changed are things most salon owners completely ignore.\n\nIf anyone wants the breakdown, let me know in the comments."'
)

add_script_block(
    "Post 4 — The $48K Question:",
    '"Did a quick calculation the other day.\n\nIf you\'re spending $3k/month on ads, that\'s $36k/year. Add in the booking software fees, the social media manager, and the random \'boost this post\' spending \u2014 you\'re looking at $48k+ per year just to stay visible.\n\nWhat if you could get MORE bookings for $0/month in ad spend?\n\nThat\'s what I\'ve been helping salon owners do."'
)

add_script_block(
    "Post 5 — Referral System:",
    '"Question for salon owners: How many referrals do you get per month?\n\nNot just \'word of mouth\' but actual referrals from a system?\n\nMost salon owners I talk to say \'a few\' but don\'t have a real system. The ones who do are getting 5-10 per month consistently.\n\nIt\'s one of the 4 pillars I teach. Happy to share the basics if anyone\'s interested."'
)

add_script_block(
    "Comment CTA (when people respond to your posts):",
    '"DM me and I\'ll share the details!" — Then move the conversation to DM and use the Instagram DM follow-up scripts adapted for Messenger.'
)

doc.add_page_break()

# ===================== PART 4: GHL EMAIL SEQUENCES =====================
add_heading("PART 4: COLD EMAIL SEQUENCES FOR GHL (9 Emails)", 1)
doc.add_paragraph("Load all 9 emails into GoHighLevel as an automated workflow. Use {{contact.first_name}} for personalization. A/B test subject lines.")

p = doc.add_paragraph()
r = p.add_run("Sequence Map:")
r.bold = True
doc.add_paragraph("Phase 2 (Interest): Emails 1, 2, 3 \u2014 Pain, Proof, Tease (Days 1, 3, 6)")
doc.add_paragraph("Phase 3 (Desire): Emails 4, 5, 6 \u2014 Offer Reveal, Bonus Stack, Testimonials (Days 1, 3, 6)")
doc.add_paragraph("Phase 4 (Action): Emails 7, 8, 9 \u2014 48hr, 24hr, 6hr before close")

add_heading("Phase 2 \u2014 Interest Emails", 2)

add_email_block(1, "The Pain", "Phase 2, Day 1",
    "Are you renting clients or building a business?",
    "The $48,000 mistake most salon owners make every year",
    "Most salon owners don't realize the trap they're in until they see this number...",
    """Hey {{contact.first_name}},

Quick math for you.

If you're spending $3,000/month on ads for your salon \u2014 which is on the low end for most US salons \u2014 that's $36,000 a year.

Now ask yourself this: if you stopped paying tomorrow, would your bookings dry up?

If the answer is yes, you're not building a business. You're renting clients.

I've talked to hundreds of salon owners who've been stuck on this treadmill for years. The worst part? The moment the ads stop, everything stops. There's no system underneath. No loyalty. No referrals. Just silence.

The good news: there is a way out. And it doesn't require more ad spend. It actually requires zero ad spend.

I'll show you what it looks like in my next email.

Talk soon,
[Your Name]""",
    "P.S. \u2014 The number most salon owners discover when they finally do the math? $48,000+/year gone. Just in ads. With inconsistent results."
)

add_email_block(2, "Social Proof", "Phase 2, Day 3",
    "23 clients came back. Zero ad spend.",
    "How Fatima stopped paying $4,000/month for ads (and got MORE bookings)",
    "Maria thought those clients were gone forever. Then she ran one simple campaign...",
    """Hey {{contact.first_name}},

Yesterday I talked about the ad treadmill \u2014 spending $3k\u2013$8k/month just to stay visible.

Today I want to show you what happens when salon owners get off it.

Meet Fatima Al-Rashid:
She was spending $4,000/month on Instagram ads. Barely any return. After implementing just 3 tactics from a system I'll share with you, she got more bookings than ever \u2014 and completely stopped paid ads.
Current ad spend: $0. Extra bookings per month: +35.

Meet Maria Petrova:
The reactivation campaign brought back 23 clients she thought were gone forever. Using simple, non-pushy WhatsApp scripts. Week 1. The course paid for itself immediately.

These aren't outliers. This is what happens when you build a real marketing system instead of renting traffic.

The system has 4 pillars:
1. RANK \u2014 dominate Google Maps so clients find YOU first
2. CONVERT \u2014 Instagram content that turns followers into bookings
3. REACTIVATE \u2014 bring back 30\u201345% of clients who haven't visited in 90+ days
4. REFER \u2014 simple system that generates 5\u201310 referrals per month

Tomorrow I'll show you exactly what's inside.

[Your Name]"""
)

add_email_block(3, "The Tease", "Phase 2, Day 6",
    "The 8-module system that filled 500+ salon calendars (sneak peek)",
    "Before I show you the price \u2014 here's what's inside",
    "Over the last two days I introduced you to the problem and the proof. Today, the system.",
    """Hey {{contact.first_name}},

Over the last two emails I introduced you to the problem (the ad treadmill) and the proof (Fatima, Maria, and 500+ others).

Today I want to give you a peek at the actual system.

The Salon Marketing Masterclass is 8 video modules \u2014 5+ hours of step-by-step training specifically built for salon owners. No fluff. No theory. Just tactics you can implement the same week.

Here's what's inside:
\u2022 Module 1 \u2014 Foundation & Quick Wins Audit
\u2022 Module 2 \u2014 Google Business Domination (rank #1 for 'salon near me')
\u2022 Module 3 \u2014 Instagram Content That Converts (30-day calendar included)
\u2022 Module 4 \u2014 Before/After Content System
\u2022 Module 5 \u2014 Client Reactivation Campaign (the 3-wave system)
\u2022 Module 6 \u2014 Referral & Partnership Strategy
\u2022 Module 7 \u2014 Promotions & QR Capture
\u2022 Module 8 \u2014 Review Amplification (turn every 5-star review into 7 assets)

Plus 6 implementation toolkits \u2014 templates, scripts, and checklists worth $544 on their own.

In my next email, I'll send you the full offer and the price.

Spoiler: it's not $497.

[Your Name]"""
)

add_heading("Phase 3 \u2014 Desire Emails", 2)

add_email_block(4, "Offer Reveal", "Phase 3, Day 1",
    "Here's everything you get \u2014 and the price drops today",
    "$1,041 in value. $97 today. Here's the full breakdown.",
    "The full Salon Marketing Masterclass is open. Here's what's included...",
    """Hey {{contact.first_name}},

I promised you the full offer today. Here it is.

The Salon Marketing Masterclass includes:
\u2022 8 Video Modules (5+ hours of step-by-step training) \u2014 $497 value
\u2022 30-Day Action Planner \u2014 $89 value
\u2022 GBP Domination Kit (100+ photo ideas, templates, schedule) \u2014 $79 value
\u2022 Instagram Content Vault (90 days of captions + Canva templates) \u2014 $99 value
\u2022 Reactivation Swipe File (copy-paste WhatsApp + SMS scripts) \u2014 $99 value
\u2022 12-Month Seasonal Calendar \u2014 $79 value
\u2022 Review Amplification Pack \u2014 $99 value

Total value: $1,041.

Launch price today: $97.

One-time payment. Lifetime access. 30-day money-back guarantee \u2014 implement the tactics, don't see results, email us with proof and you get every dollar back.

Enroll now: evolvxai.com/salon01-3329

This price does not last. I'll tell you exactly when it ends in my next email.

[Your Name]""",
    "P.S. \u2014 500+ salon owners have already gone through this. The only common thing among the ones who didn't see results? They never implemented."
)

add_email_block(5, "Bonus Stack", "Phase 3, Day 3",
    "Here's what you lose if you wait (the bonuses expire too)",
    "The bonuses disappear before the price goes up",
    "It's not just the price that has a deadline...",
    """Hey {{contact.first_name}},

Yesterday I showed you the full offer \u2014 $1,041 in training and toolkits for $97.

Today I want to make sure you know: it's not just the price that has a deadline.

The launch bonuses \u2014 the Reactivation Swipe File, the Instagram Content Vault, the 30-Day Action Planner \u2014 these are included at $97 only for the launch window. After that, the course is still available, but the bonuses stack goes away.

Why does this matter?

Because the Reactivation Swipe File alone is the asset that brought Maria 23 clients back in Week 1. That's not a nice-to-have. That's direct revenue the day you implement it.

The Instagram Content Vault gives you 90 days of captions and Canva templates \u2014 so you never stare at a blank post again.

The 30-Day Action Planner tells you exactly what to do every morning. No guessing.

All of this disappears when the launch window closes.

Enroll before that happens: evolvxai.com/salon01-3329

[Your Name]"""
)

add_email_block(6, "Testimonials", "Phase 3, Day 6",
    "Real salon owners. Real numbers. Read these.",
    "From $4,000/month in ads to $0 \u2014 their stories",
    "Before you decide, hear from the people who've already done this...",
    """Hey {{contact.first_name}},

I'm going to keep this one short and let the results do the talking.

\u2014\u2014\u2014

"I was spending $4,000/month on Instagram ads with barely any return. After implementing just 3 tactics from this course, I'm getting more bookings than ever and I've completely stopped paid ads."
\u2014 Fatima Al-Rashid  |  Ad spend now: $0  |  Extra bookings: +35/month

\u2014\u2014\u2014

"The reactivation campaign brought back 23 clients I thought were gone forever. The WhatsApp scripts are gold. Simple, not pushy, and they actually work."
\u2014 Maria Petrova  |  23 clients reactivated  |  Paid for itself in Week 1

\u2014\u2014\u2014

500+ salon owners have enrolled. These aren't the exceptions \u2014 they're what happens when you actually implement.

You're 30 days away from results like these. The only thing standing between you and a full calendar is the decision.

Enroll here: evolvxai.com/salon01-3329

[Your Name]""",
    "P.S. \u2014 The launch price closes soon. Don't be the person who checks this email after it's gone and wishes they'd moved faster."
)

add_heading("Phase 4 \u2014 Action Emails (Last Chance)", 2)

add_email_block(7, "48-Hour Warning", "Phase 4, Day 1",
    "48 hours left \u2014 then the price goes back to $497",
    "The math on waiting vs. enrolling right now",
    "The launch window is closing. Here's what you need to know...",
    """Hey {{contact.first_name}},

Quick heads up: the $97 launch price for Salon Marketing Masterclass closes in 48 hours.

After that, it goes back to $497.

I'm not saying this to pressure you. I'm saying it because I've seen what happens when salon owners wait \u2014 they end up spending another $4,000\u2013$8,000 on ads next month, and then the month after that, while this system sits one click away.

What you get for $97 today:
\u2022 8 modules \u2014 5+ hours of implementation-ready training
\u2022 6 bonus toolkits \u2014 templates, scripts, and calendars
\u2022 30-day money-back guarantee \u2014 zero risk
\u2022 Lifetime access with free updates

Enroll now before the window closes: evolvxai.com/salon01-3329

[Your Name]"""
)

add_email_block(8, "24-Hour Warning", "Phase 4, Day 2",
    "Last chance \u2014 24 hours before the price changes",
    "What if it doesn't work? (Read this guarantee)",
    "This is the final reminder. After midnight tomorrow, the $97 price is gone...",
    """Hey {{contact.first_name}},

24 hours.

That's how long the $97 price is still available for Salon Marketing Masterclass.

Tomorrow it goes back to $497.

I want to answer the most common objection I hear at this point:

"What if it doesn't work for my salon?"

That's exactly why there's a 30-day money-back guarantee. Implement the tactics for 30 days. If you don't see a measurable increase in bookings, email us with proof you did the work \u2014 you get a full refund. No questions asked.

There is genuinely no risk here. The only risk is spending another month paying for ads while the system that replaces them costs $97.

Last chance: evolvxai.com/salon01-3329

[Your Name]""",
    "P.S. \u2014 500+ salon owners already made this decision. You know what they all have in common now? Full calendars."
)

add_email_block(9, "Final Close", "Phase 4, Day 3 \u2014 6hrs before",
    "This is it \u2014 closing in 6 hours [FINAL NOTICE]",
    "Final email. Then it's done.",
    "I'm closing the cart in 6 hours. This is the last email I'll send about this...",
    """Hey {{contact.first_name}},

I'm closing the $97 launch price in 6 hours.

This is the last email I'll send about this. After tonight, the price goes to $497 and the launch bonuses are gone.

If you've been on the fence \u2014 this is the moment.

I don't believe in manufactured urgency. This deadline is real. The system works. The guarantee removes all risk.

If you want to stop bleeding money on ads and build a salon calendar that fills itself organically \u2014 this is how.

Final link: evolvxai.com/salon01-3329

If you're not ready, no hard feelings. But if you are \u2014 move now.

[Your Name]""",
    "P.S. \u2014 After the cart closes tonight, I'll be back with new content and new topics. But this offer \u2014 $97, full bonuses, this guarantee \u2014 won't come back at this price. This is it."
)

doc.add_page_break()

# ===================== PART 5: WHATSAPP SCRIPTS =====================
add_heading("PART 5: WHATSAPP FOLLOW-UP SCRIPTS", 1)
doc.add_paragraph("Use voice notes wherever possible \u2014 they convert significantly better than text because they feel personal. Only use these in Phase 3+ with warm leads who have already engaged.")

add_script_block(
    "Voice Note 1 \u2014 Warm Lead Follow-Up:",
    '"Hey [Name], just wanted to follow up personally \u2014 I know we spoke about your salon marketing a few weeks back. The Salon Marketing Masterclass is closing the launch price on [date] and I wanted to give you a heads up before I pull it. It\'s $97 right now, goes back to $497 after that. Happy to answer any questions before you decide \u2014 just reply here."'
)

add_script_block(
    "Voice Note 2 \u2014 Post-Chat Follow-Up:",
    '"Hey [Name], it\'s [Your Name]. I wanted to check in after our chat. Did the 4-pillar system make sense for your salon? The launch price closes [date] \u2014 just wanted to make sure you had all the info you needed before deciding. No pressure at all."'
)

add_script_block(
    "Text Message 1 \u2014 Link Drop:",
    '"Hey [Name]! Here\'s the link to the Salon Marketing Masterclass we talked about: evolvxai.com/salon01-3329 \u2014 $97 (normally $497), 30-day guarantee. Let me know if you have any questions!"'
)

add_script_block(
    "Text Message 2 \u2014 Final Reminder:",
    '"Hey [Name] \u2014 quick heads up: the $97 price closes tonight. After that it\'s $497. Just didn\'t want you to miss it. Link: evolvxai.com/salon01-3329"'
)

doc.add_page_break()

# ===================== PART 6: THREADS SCRIPTS =====================
add_heading("PART 6: THREADS / X OUTREACH SCRIPTS", 1)

add_heading("Reply Scripts (Engage on Salon Owner Posts)", 2)

add_script_block(
    "Reply A \u2014 On ad complaints:",
    '"This is so true \u2014 I\'ve been working with salon owners on exactly this. The ones who switched to organic marketing are spending $0 on ads and getting more bookings than before."'
)

add_script_block(
    "Reply B \u2014 On marketing questions:",
    '"The 3 things I\'ve seen work best: Google Maps optimization, client reactivation (WhatsApp scripts work incredibly well), and a simple referral system. No ads needed."'
)

add_heading("Original Post Scripts", 2)

add_script_block(
    "Post A:",
    '"Hot take: salon owners spending $3k-$8k/month on ads are renting their clients, not building a business. The smart ones are building organic systems. I just helped 500+ do exactly that."'
)

add_script_block(
    "Post B:",
    '"A salon owner sent 3 WhatsApp messages to clients who hadn\'t visited in 90+ days. 23 came back in a week. $0 spent. The right message > the biggest ad budget."'
)

add_script_block(
    "Post C:",
    '"The $48,000 question: If you\'re a salon owner spending $3k/month on ads, that\'s $36k/year. Add everything else and you\'re at $48k+ just to stay visible. What if you could get MORE bookings for $0?"'
)

# SAVE
path = "/Users/nishamore/Documents/Salon world wide/Complete_Outreach_Scripts.docx"
doc.save(path)
print(f"Complete Outreach Scripts saved to: {path}")
print(f"File size: {os.path.getsize(path)} bytes")
