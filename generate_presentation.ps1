# PowerShell script to generate a 10-slide PowerPoint presentation with custom design and speaker notes
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function RGB($r, $g, $b) {
    return [int]($r + ($g * 256) + ($b * 65536))
}

$BG_COLOR       = RGB 15 23 42      # #0F172A Dark Slate/Navy
$CARD_COLOR     = RGB 30 41 59      # #1E293B Slate Card
$CARD_BORDER    = RGB 51 65 85      # #334155 Border
$CARD_ACCENT    = RGB 30 58 138     # #1E3A8A Blue tint
$TEXT_WHITE     = RGB 255 255 255  # #FFFFFF
$TEXT_MUTED     = RGB 148 163 184  # #94A3B8
$TEXT_ACCENT    = RGB 56 189 248   # #38BDF8 Sky Cyan
$TEXT_GOLD      = RGB 245 158 11   # #F59E0B Amber
$TEXT_GREEN     = RGB 52 211 153   # #34D399 Emerald
$TEXT_PURPLE    = RGB 192 132 252  # #C084FC Lavender
$TEXT_ROSE      = RGB 244 114 182  # #F472B6 Pink/Rose

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoFalse)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

function Add-SlideBackground($slide) {
    $bg = $slide.Shapes.AddShape(1, 0, 0, 960, 540)
    $bg.Fill.Solid()
    $bg.Fill.ForeColor.RGB = $BG_COLOR
    $bg.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    
    $accentLine = $slide.Shapes.AddShape(1, 0, 0, 960, 4)
    $accentLine.Fill.Solid()
    $accentLine.Fill.ForeColor.RGB = $TEXT_ACCENT
    $accentLine.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
}

function Add-Header($slide, $badgeText, $titleText, $subtitleText) {
    $badge = $slide.Shapes.AddShape(5, 50, 24, 180, 22)
    $badge.Fill.Solid()
    $badge.Fill.ForeColor.RGB = RGB 30 58 138
    $badge.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $badge.TextFrame.TextRange.Text = $badgeText.ToUpper()
    $badge.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $badge.TextFrame.TextRange.Font.Size = 9
    $badge.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $badge.TextFrame.TextRange.Font.Color.RGB = $TEXT_ACCENT
    $badge.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    
    $title = $slide.Shapes.AddTextbox(1, 50, 48, 860, 38)
    $title.TextFrame.TextRange.Text = $titleText
    $title.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $title.TextFrame.TextRange.Font.Size = 22
    $title.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $title.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE
    $title.TextFrame.MarginLeft = 0
    $title.TextFrame.MarginTop = 0
    
    if ($subtitleText) {
        $sub = $slide.Shapes.AddTextbox(1, 50, 84, 860, 24)
        $sub.TextFrame.TextRange.Text = $subtitleText
        $sub.TextFrame.TextRange.Font.Name = 'Segoe UI'
        $sub.TextFrame.TextRange.Font.Size = 11
        $sub.TextFrame.TextRange.Font.Color.RGB = $TEXT_MUTED
        $sub.TextFrame.MarginLeft = 0
        $sub.TextFrame.MarginTop = 0
    }
}

function Set-SpeakerNote($slide, $noteText) {
    try {
        $note = $slide.NotesPage.Shapes.Placeholders.Item(2)
        $note.TextFrame.TextRange.Text = $noteText
    } catch {
        # ignore note errors if placeholder differs
    }
}

# -------------------------------------------------------------
# SLIDE 1: TITLE SLIDE
# -------------------------------------------------------------
$s1 = $pres.Slides.Add(1, 12)
Add-SlideBackground $s1

$glow = $s1.Shapes.AddShape(9, 230, 70, 500, 400)
$glow.Fill.Solid()
$glow.Fill.ForeColor.RGB = RGB 23 37 84
$glow.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

$b1 = $s1.Shapes.AddShape(5, 340, 50, 280, 28)
$b1.Fill.Solid()
$b1.Fill.ForeColor.RGB = RGB 30 58 138
$b1.Line.ForeColor.RGB = $TEXT_ACCENT
$b1.Line.Weight = 1.2
$b1.TextFrame.TextRange.Text = 'GRADE 11  *  BSB-1 PROJECT WORK'
$b1.TextFrame.TextRange.Font.Name = 'Segoe UI'
$b1.TextFrame.TextRange.Font.Size = 11
$b1.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$b1.TextFrame.TextRange.Font.Color.RGB = $TEXT_ACCENT
$b1.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$t1 = $s1.Shapes.AddTextbox(1, 80, 95, 800, 75)
$t1.TextFrame.TextRange.Text = 'BE A BETTER FRIEND!'
$t1.TextFrame.TextRange.Font.Name = 'Segoe UI'
$t1.TextFrame.TextRange.Font.Size = 44
$t1.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$t1.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE
$t1.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$sub1 = $s1.Shapes.AddTextbox(1, 100, 175, 760, 35)
$sub1.TextFrame.TextRange.Text = 'A Friendship Campaign for a Kinder, Safer and More Supportive School'
$sub1.TextFrame.TextRange.Font.Name = 'Segoe UI'
$sub1.TextFrame.TextRange.Font.Size = 15
$sub1.TextFrame.TextRange.Font.Color.RGB = $TEXT_MUTED
$sub1.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$scard = $s1.Shapes.AddShape(5, 160, 225, 640, 95)
$scard.Fill.Solid()
$scard.Fill.ForeColor.RGB = $CARD_COLOR
$scard.Line.ForeColor.RGB = $TEXT_GOLD
$scard.Line.Weight = 2

$slabel = $s1.Shapes.AddTextbox(1, 180, 235, 600, 25)
$slabel.TextFrame.TextRange.Text = 'OFFICIAL CAMPAIGN SLOGAN'
$slabel.TextFrame.TextRange.Font.Name = 'Segoe UI'
$slabel.TextFrame.TextRange.Font.Size = 11
$slabel.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$slabel.TextFrame.TextRange.Font.Color.RGB = $TEXT_GOLD
$slabel.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$stext = $s1.Shapes.AddTextbox(1, 180, 260, 600, 50)
$stext.TextFrame.TextRange.Text = '"Be the friend you wish to have!"'
$stext.TextFrame.TextRange.Font.Name = 'Segoe UI'
$stext.TextFrame.TextRange.Font.Size = 24
$stext.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$stext.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE
$stext.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$teamLabels = @(
    "Speaker 1: Campaign Leader`r`nOpening Speech & Mission",
    "Speaker 2: Rules Specialist`r`n5 Golden Rules & Examples",
    "Speaker 3: Advice & Grammar`r`n3 Tips & Language Rules",
    "Speaker 4: Action & Q&A`r`nChallenge, Poster & Q&A"
)
for ($i = 0; $i -lt 4; $i++) {
    $tx = 60 + ($i * 215)
    $tbox = $s1.Shapes.AddShape(5, $tx, 350, 200, 85)
    $tbox.Fill.Solid()
    $tbox.Fill.ForeColor.RGB = RGB 24 33 51
    $tbox.Line.ForeColor.RGB = $CARD_BORDER
    $tbox.TextFrame.TextRange.Text = "MEMBER $($i+1)`r`n" + $teamLabels[$i]
    $tbox.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $tbox.TextFrame.TextRange.Font.Size = 10
    $tbox.TextFrame.TextRange.Font.Color.RGB = $TEXT_MUTED
    $tbox.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $tbox.TextFrame.TextRange.Lines(1).Font.Size = 11
    $tbox.TextFrame.TextRange.Lines(1).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $tbox.TextFrame.TextRange.Lines(1).Font.Color.RGB = $TEXT_ACCENT
}

Set-SpeakerNote $s1 @"
[SPEAKER 1 - OPENING SLIDE]
Good morning, dear teachers and classmates! Today, our 11th-grade group is proud to present our BSB-1 Project Work: 'BE A BETTER FRIEND!'
Our official campaign slogan is: 'Be the friend you wish to have!'
Our team consists of 4 members, each covering crucial aspects of friendship, rules, practical advice, and interactive activities.
Let's begin!
"@

# -------------------------------------------------------------
# SLIDE 2: OUR MISSION & CORE VALUES
# -------------------------------------------------------------
$s2 = $pres.Slides.Add(2, 12)
Add-SlideBackground $s2
Add-Header $s2 'Section 01' 'Our Campaign Mission' 'Why friendship matters and how our campaign makes a positive difference'

$mcard = $s2.Shapes.AddShape(5, 50, 115, 860, 95)
$mcard.Fill.Solid()
$mcard.Fill.ForeColor.RGB = $CARD_COLOR
$mcard.Line.ForeColor.RGB = $TEXT_ACCENT
$mcard.Line.Weight = 1.5

$mt1 = $s2.Shapes.AddTextbox(1, 70, 125, 820, 75)
$mt1.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
$mt1.TextFrame.TextRange.Text = '"Our campaign encourages students to build stronger friendships through kindness, trust, honesty, respect, and support. A good friendship is not about being perfect; it is about being there for each other, especially when life becomes difficult."'
$mt1.TextFrame.TextRange.Font.Name = 'Segoe UI'
$mt1.TextFrame.TextRange.Font.Size = 13.5
$mt1.TextFrame.TextRange.Font.Italic = [Microsoft.Office.Core.MsoTriState]::msoTrue
$mt1.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE

$pHeader = $s2.Shapes.AddTextbox(1, 50, 225, 860, 25)
$pHeader.TextFrame.TextRange.Text = 'THE 5 CORE PILLARS OF OUR CAMPAIGN'
$pHeader.TextFrame.TextRange.Font.Name = 'Segoe UI'
$pHeader.TextFrame.TextRange.Font.Size = 12
$pHeader.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pHeader.TextFrame.TextRange.Font.Color.RGB = $TEXT_ACCENT

$pillars = @(
    @{ Title = 'Kindness'; Sub = 'Mehribonlik'; Desc = 'Small daily gestures that make classmates feel valued and safe.'; Color = $TEXT_GOLD },
    @{ Title = 'Trust'; Sub = 'Ishonch'; Desc = 'Keeping secrets safe and being someone friends can always rely on.'; Color = $TEXT_ACCENT },
    @{ Title = 'Honesty'; Sub = 'Haqiqatgoylik'; Desc = 'Speaking truth kindly without gossip, assumptions, or rumors.'; Color = $TEXT_GREEN },
    @{ Title = 'Respect'; Sub = 'Hurmat'; Desc = 'Valuing different opinions, personal boundaries, and unique feelings.'; Color = $TEXT_PURPLE },
    @{ Title = 'Support'; Sub = 'Qollab-quvvatlash'; Desc = 'Standing beside friends when life gets tough or stressful.'; Color = $TEXT_ROSE }
)

for ($i = 0; $i -lt 5; $i++) {
    $px = 50 + ($i * 176)
    $p = $pillars[$i]
    
    $pc = $s2.Shapes.AddShape(5, $px, 260, 166, 230)
    $pc.Fill.Solid()
    $pc.Fill.ForeColor.RGB = $CARD_COLOR
    $pc.Line.ForeColor.RGB = $p.Color
    $pc.Line.Weight = 1.2
    
    $pt = $s2.Shapes.AddTextbox(1, $px + 10, 270, 146, 210)
    $pt.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pt.TextFrame.TextRange.Text = "0$($i+1)`r`n$($p.Title)`r`n$($p.Sub)`r`n`r`n$($p.Desc)"
    $pt.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $pt.TextFrame.TextRange.Font.Size = 11
    $pt.TextFrame.TextRange.Font.Color.RGB = $TEXT_MUTED
    $pt.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    
    $pt.TextFrame.TextRange.Lines(1).Font.Size = 14
    $pt.TextFrame.TextRange.Lines(1).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pt.TextFrame.TextRange.Lines(1).Font.Color.RGB = $p.Color
    
    $pt.TextFrame.TextRange.Lines(2).Font.Size = 14
    $pt.TextFrame.TextRange.Lines(2).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pt.TextFrame.TextRange.Lines(2).Font.Color.RGB = $TEXT_WHITE
}

Set-SpeakerNote $s2 @"
[SPEAKER 1 - MISSION & PILLARS]
What is the core mission of our campaign?
It is not about expecting perfection from our friends. Real friendship means standing together when life becomes difficult.
Our campaign is built on five strong pillars:
1. Kindness - making everyone feel welcome.
2. Trust - maintaining confidentiality.
3. Honesty - clear, open communication.
4. Respect - valuing differences.
5. Support - helping through struggles.
"@

# -------------------------------------------------------------
# SLIDE 3: FIVE RULES OF FRIENDSHIP
# -------------------------------------------------------------
$s3 = $pres.Slides.Add(3, 12)
Add-SlideBackground $s3
Add-Header $s3 'Section 02' 'Five Golden Rules of Friendship' 'Essential rules to maintain healthy, resilient, and honest connections'

$rules = @(
    @{ Title = '1. Listen Carefully'; Text = 'Give your friend time to speak without rushing them. Do not interrupt or judge too quickly.'; Accent = $TEXT_ACCENT },
    @{ Title = '2. Be Honest and Trustworthy'; Text = 'Tell the truth and keep private information confidential. Trust is easy to lose and hard to rebuild.'; Accent = $TEXT_GOLD },
    @{ Title = '3. Show Respect'; Text = 'Respect different opinions, unique interests, personal boundaries, and sensitive feelings.'; Accent = $TEXT_GREEN },
    @{ Title = '4. Offer Support'; Text = 'If your friend has a problem, ask how you can help instead of ignoring them or walking away.'; Accent = $TEXT_PURPLE },
    @{ Title = '5. Solve Problems Calmly'; Text = 'Talk about misunderstandings politely. Apologise sincerely when you are wrong and forgive when you can.'; Accent = $TEXT_ROSE }
)

for ($i = 0; $i -lt 5; $i++) {
    $ry = 115 + ($i * 77)
    $r = $rules[$i]
    
    $rc = $s3.Shapes.AddShape(5, 50, $ry, 860, 68)
    $rc.Fill.Solid()
    $rc.Fill.ForeColor.RGB = $CARD_COLOR
    $rc.Line.ForeColor.RGB = $CARD_BORDER
    $rc.Line.Weight = 1
    
    $nb = $s3.Shapes.AddShape(5, 62, $ry + 12, 44, 44)
    $nb.Fill.Solid()
    $nb.Fill.ForeColor.RGB = RGB 24 33 51
    $nb.Line.ForeColor.RGB = $r.Accent
    $nb.Line.Weight = 1.5
    $nb.TextFrame.TextRange.Text = "$($i+1)"
    $nb.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $nb.TextFrame.TextRange.Font.Size = 18
    $nb.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $nb.TextFrame.TextRange.Font.Color.RGB = $r.Accent
    $nb.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    
    $rt = $s3.Shapes.AddTextbox(1, 120, $ry + 10, 770, 50)
    $rt.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $rt.TextFrame.TextRange.Text = "$($r.Title)`r`n$($r.Text)"
    $rt.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $rt.TextFrame.TextRange.Font.Size = 11
    $rt.TextFrame.TextRange.Font.Color.RGB = $TEXT_MUTED
    
    $rt.TextFrame.TextRange.Lines(1).Font.Size = 13.5
    $rt.TextFrame.TextRange.Lines(1).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $rt.TextFrame.TextRange.Lines(1).Font.Color.RGB = $TEXT_WHITE
}

Set-SpeakerNote $s3 @"
[SPEAKER 2 - THE FIVE RULES]
Now, let's explore the Five Golden Rules of Friendship:
First, listen carefully. Listening is better than interrupting.
Second, be honest and trustworthy. Private secrets must remain confidential.
Third, show respect even when friends hold different opinions.
Fourth, offer support; ask 'How can I help you today?'
Fifth, solve problems calmly. Apologising when you are wrong is a sign of strength, not weakness!
"@

# -------------------------------------------------------------
# SLIDE 4: THREE PIECES OF ADVICE
# -------------------------------------------------------------
$s4 = $pres.Slides.Add(4, 12)
Add-SlideBackground $s4
Add-Header $s4 'Section 03' 'Three Pieces of Practical Advice' 'Actionable guidance to turn friendship values into everyday classroom practice'

$advices = @(
    @{
        Tag = 'EMPATHY & LISTENING'
        Title = 'When a Friend Feels Sad'
        Quote = '"If your friend feels sad, listen to them and ask what they need."'
        Detail = 'Do not rush to give unsolicited advice or solve everything immediately. Sitting quietly and listening without judgment makes them feel safe and respected.'
        Grammar = 'Grammar: First Conditional (If + Present, Imperative)'
        Color = $TEXT_ACCENT
    },
    @{
        Tag = 'HONESTY & INTEGRITY'
        Title = 'Handling Misunderstandings'
        Quote = '"You should talk honestly instead of spreading rumours or making assumptions."'
        Detail = 'Gossip breaks classroom unity. Speak openly and politely face-to-face with your friend rather than whispering behind their back or assuming bad intentions.'
        Grammar = 'Grammar: Modal Verb "should" for strong advice'
        Color = $TEXT_GOLD
    },
    @{
        Tag = 'INCLUSION & KINDNESS'
        Title = 'Welcoming Lonely Classmates'
        Quote = '"You could invite a lonely classmate to join your group and help them feel welcome."'
        Detail = 'No one deserves to feel left out or eat lunch in isolation. A simple friendly invitation can transform someone’s school life and create lasting bonds.'
        Grammar = 'Grammar: Modal Verb "could" for polite suggestions'
        Color = $TEXT_GREEN
    }
)

for ($i = 0; $i -lt 3; $i++) {
    $ax = 50 + ($i * 295)
    $a = $advices[$i]
    
    $ac = $s4.Shapes.AddShape(5, $ax, 115, 275, 390)
    $ac.Fill.Solid()
    $ac.Fill.ForeColor.RGB = $CARD_COLOR
    $ac.Line.ForeColor.RGB = $a.Color
    $ac.Line.Weight = 1.5
    
    $atag = $s4.Shapes.AddShape(5, $ax + 15, 128, 245, 22)
    $atag.Fill.Solid()
    $atag.Fill.ForeColor.RGB = RGB 24 33 51
    $atag.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $atag.TextFrame.TextRange.Text = $a.Tag
    $atag.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $atag.TextFrame.TextRange.Font.Size = 9
    $atag.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $atag.TextFrame.TextRange.Font.Color.RGB = $a.Color
    $atag.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    
    $atitle = $s4.Shapes.AddTextbox(1, $ax + 15, 155, 245, 30)
    $atitle.TextFrame.TextRange.Text = $a.Title
    $atitle.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $atitle.TextFrame.TextRange.Font.Size = 13
    $atitle.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $atitle.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE
    $atitle.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    
    $aq = $s4.Shapes.AddShape(5, $ax + 15, 190, 245, 75)
    $aq.Fill.Solid()
    $aq.Fill.ForeColor.RGB = RGB 24 33 51
    $aq.Line.ForeColor.RGB = $a.Color
    $aq.Line.Weight = 1
    $aq.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $aq.TextFrame.TextRange.Text = $a.Quote
    $aq.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $aq.TextFrame.TextRange.Font.Size = 10.5
    $aq.TextFrame.TextRange.Font.Italic = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $aq.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE
    $aq.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    
    $ad = $s4.Shapes.AddTextbox(1, $ax + 15, 275, 245, 215)
    $ad.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $ad.TextFrame.TextRange.Text = "$($a.Detail)`r`n`r`n$($a.Grammar)"
    $ad.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $ad.TextFrame.TextRange.Font.Size = 10.5
    $ad.TextFrame.TextRange.Font.Color.RGB = $TEXT_MUTED
}

Set-SpeakerNote $s4 @"
[SPEAKER 3 - THREE PIECES OF ADVICE]
Here are three practical pieces of advice for our daily school life:
1. When your friend is feeling down: Just listen without judging.
2. In times of disagreement: You should speak honestly and avoid gossip.
3. When someone is left out: You could invite them to your table.
Notice our grammar structures: conditionals ('if...') and modal verbs ('should', 'could').
"@

# -------------------------------------------------------------
# SLIDE 5: SHORT CAMPAIGN SPEECH (4 SPEAKERS)
# -------------------------------------------------------------
$s5 = $pres.Slides.Add(5, 12)
Add-SlideBackground $s5
Add-Header $s5 'Section 04' 'Campaign Speech (4 Speakers)' 'Balanced oral presentation delivery plan for all four group members'

$speakers = @(
    @{
        Role = 'SPEAKER 1 - INTRODUCTION'
        Duty = 'Campaign Leader & Slogan'
        Text = '"Good morning, everyone! Today, we are presenting our campaign, Be a Better Friend! Friendship makes our lives happier and our challenges easier. However, strong friendships need time, trust, and effort. Our slogan is: Be the friend you wish to have!"'
        Color = $TEXT_ACCENT
    },
    @{
        Role = 'SPEAKER 2 - THE FIVE RULES'
        Duty = 'Rules Specialist & Core Values'
        Text = '"A good friend listens carefully, tells the truth, shows respect, offers support, and solves problems calmly. For example, listening is better than interrupting, and honest communication is stronger than gossip. We should respect our friends even when we have different opinions."'
        Color = $TEXT_GOLD
    },
    @{
        Role = 'SPEAKER 3 - ADVICE & GRAMMAR'
        Duty = 'Language Specialist & Tips'
        Text = '"Here are three simple tips. If your friend feels sad, listen without judging. You should speak honestly and never spread rumours. You could also invite someone who feels left out to join you. If we support one another, our classroom will become a kinder place."'
        Color = $TEXT_GREEN
    },
    @{
        Role = 'SPEAKER 4 - CALL TO ACTION'
        Duty = 'Conclusion & Challenge Leader'
        Text = '"Let us take a friendship challenge! This week, do one kind thing every day: help someone, give a sincere compliment, or include a classmate in an activity. Small actions create stronger friendships. Remember, a real friend does not have to solve every problem; sometimes, being present is enough."'
        Color = $TEXT_PURPLE
    }
)

for ($i = 0; $i -lt 4; $i++) {
    $row = [Math]::Floor($i / 2)
    $col = $i % 2
    $sx = 50 + ($col * 440)
    $sy = 115 + ($row * 195)
    $sp = $speakers[$i]
    
    $sc = $s5.Shapes.AddShape(5, $sx, $sy, 420, 180)
    $sc.Fill.Solid()
    $sc.Fill.ForeColor.RGB = $CARD_COLOR
    $sc.Line.ForeColor.RGB = $sp.Color
    $sc.Line.Weight = 1.2
    
    $sh = $s5.Shapes.AddTextbox(1, $sx + 15, $sy + 10, 390, 26)
    $sh.TextFrame.TextRange.Text = "$($sp.Role)  *  $($sp.Duty)"
    $sh.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $sh.TextFrame.TextRange.Font.Size = 10.5
    $sh.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $sh.TextFrame.TextRange.Font.Color.RGB = $sp.Color
    
    $st = $s5.Shapes.AddTextbox(1, $sx + 15, $sy + 36, 390, 130)
    $st.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $st.TextFrame.TextRange.Text = $sp.Text
    $st.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $st.TextFrame.TextRange.Font.Size = 11
    $st.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE
}

Set-SpeakerNote $s5 @"
[ALL SPEAKERS - ROTATION]
This slide serves as the script overview.
Each speaker takes their turn smoothly:
- Speaker 1 introduces the theme and sets an inspiring tone.
- Speaker 2 explains the 5 rules clearly.
- Speaker 3 delivers the advice and language highlights.
- Speaker 4 motivates the room with the call to action!
"@

# -------------------------------------------------------------
# SLIDE 6: LANGUAGE FEATURES & GRAMMAR FOCUS
# -------------------------------------------------------------
$s6 = $pres.Slides.Add(6, 12)
Add-SlideBackground $s6
Add-Header $s6 'Section 05' 'Language Features & Grammar Focus' 'Key grammatical structures required by Grade 11 BSB-1 evaluation criteria'

$grammar = @(
    @{
        Category = 'COMPARATIVES'
        Rule = 'Comparing actions and values (better than, stronger than, bigger)'
        Ex1 = '* "Listening is better than interrupting."'
        Ex2 = '* "Honest communication is stronger than gossip."'
        Ex3 = '* "We can make a bigger difference together than alone."'
        Color = $TEXT_ACCENT
    },
    @{
        Category = 'MODAL VERBS'
        Rule = 'Expressing advice, polite suggestion, and necessity'
        Ex1 = '* "You should speak honestly instead of spreading rumours."'
        Ex2 = '* "You could invite a lonely classmate to join your group."'
        Ex3 = '* "A real friend does not have to solve every single problem."'
        Color = $TEXT_GOLD
    },
    @{
        Category = 'CONDITIONALS'
        Rule = 'Zero & First Conditionals for cause, consequence, and advice'
        Ex1 = '* "If your friend feels sad, listen without judging."'
        Ex2 = '* "If we support one another, our classroom will become a kinder place."'
        Ex3 = '* "If both communicate honestly, friendship can become stronger."'
        Color = $TEXT_GREEN
    },
    @{
        Category = 'LINKING WORDS'
        Rule = 'Cohesion and logical transitions in discourse'
        Ex1 = '* Contrast: however, instead of, even when'
        Ex2 = '* Exemplification: for example, such as'
        Ex3 = '* Cause & Sequence: because, also, together, finally'
        Color = $TEXT_PURPLE
    }
)

for ($i = 0; $i -lt 4; $i++) {
    $row = [Math]::Floor($i / 2)
    $col = $i % 2
    $gx = 50 + ($col * 440)
    $gy = 115 + ($row * 195)
    $g = $grammar[$i]
    
    $gc = $s6.Shapes.AddShape(5, $gx, $gy, 420, 180)
    $gc.Fill.Solid()
    $gc.Fill.ForeColor.RGB = $CARD_COLOR
    $gc.Line.ForeColor.RGB = $g.Color
    $gc.Line.Weight = 1.2
    
    $gt = $s6.Shapes.AddTextbox(1, $gx + 15, $gy + 10, 390, 160)
    $gt.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $gt.TextFrame.TextRange.Text = "$($g.Category)`r`n$($g.Rule)`r`n`r`n$($g.Ex1)`r`n$($g.Ex2)`r`n$($g.Ex3)"
    $gt.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $gt.TextFrame.TextRange.Font.Size = 10.5
    $gt.TextFrame.TextRange.Font.Color.RGB = $TEXT_MUTED
    
    $gt.TextFrame.TextRange.Lines(1).Font.Size = 13
    $gt.TextFrame.TextRange.Lines(1).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $gt.TextFrame.TextRange.Lines(1).Font.Color.RGB = $g.Color
    
    $gt.TextFrame.TextRange.Lines(2).Font.Size = 10
    $gt.TextFrame.TextRange.Lines(2).Font.Color.RGB = $TEXT_WHITE
}

Set-SpeakerNote $s6 @"
[SPEAKER 3 - GRAMMAR DEMONSTRATION]
In Grade 11 BSB-1, linguistic accuracy is essential. We have integrated:
1. Comparatives: showing the contrast between good and bad habits.
2. Modals: 'should' for duty, 'could' for possibility, 'does not have to' for relief.
3. Conditionals: linking compassionate actions to positive classroom outcomes.
4. Linking words: creating natural flow and cohesion.
"@

# -------------------------------------------------------------
# SLIDE 7: INTERACTIVE CLASS ACTIVITY
# -------------------------------------------------------------
$s7 = $pres.Slides.Add(7, 12)
Add-SlideBackground $s7
Add-Header $s7 'Section 06' 'Interactive Activity: The Friendship Challenge' 'Engaging classmates and creating real behavioral change in the classroom'

$qbox = $s7.Shapes.AddShape(5, 50, 115, 860, 105)
$qbox.Fill.Solid()
$qbox.Fill.ForeColor.RGB = RGB 30 58 138
$qbox.Line.ForeColor.RGB = $TEXT_ACCENT
$qbox.Line.Weight = 2

$qtag = $s7.Shapes.AddTextbox(1, 70, 122, 820, 20)
$qtag.TextFrame.TextRange.Text = 'INTERACTIVE QUESTION FOR THE AUDIENCE:'
$qtag.TextFrame.TextRange.Font.Name = 'Segoe UI'
$qtag.TextFrame.TextRange.Font.Size = 11
$qtag.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$qtag.TextFrame.TextRange.Font.Color.RGB = $TEXT_GOLD

$qmain = $s7.Shapes.AddTextbox(1, 70, 142, 820, 65)
$qmain.TextFrame.TextRange.Text = '"What is one small thing you can do today to be a better friend?"'
$qmain.TextFrame.TextRange.Font.Name = 'Segoe UI'
$qmain.TextFrame.TextRange.Font.Size = 20
$qmain.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$qmain.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE

$steps = @(
    @{ Day = 'Day 1 - 2'; Act = 'Sincere Compliment'; Desc = 'Give a genuine compliment to a classmate you rarely talk to.'; Icon = $TEXT_ACCENT },
    @{ Day = 'Day 3 - 4'; Act = 'Help & Support'; Desc = 'Offer help with homework, carry books, or share learning materials.'; Icon = $TEXT_GOLD },
    @{ Day = 'Day 5 - 6'; Act = 'Include the Lonely'; Desc = 'Invite someone sitting alone to join your lunch table or project group.'; Icon = $TEXT_GREEN },
    @{ Day = 'Day 7'; Act = 'Gratitude & Peace'; Desc = 'Say thank you to a loyal friend, or apologise calmly after a dispute.'; Icon = $TEXT_PURPLE }
)

for ($i = 0; $i -lt 4; $i++) {
    $cx = 50 + ($i * 220)
    $st = $steps[$i]
    
    $cc = $s7.Shapes.AddShape(5, $cx, 235, 200, 220)
    $cc.Fill.Solid()
    $cc.Fill.ForeColor.RGB = $CARD_COLOR
    $cc.Line.ForeColor.RGB = $st.Icon
    $cc.Line.Weight = 1.2
    
    $db = $s7.Shapes.AddShape(5, $cx + 15, 245, 170, 24)
    $db.Fill.Solid()
    $db.Fill.ForeColor.RGB = RGB 24 33 51
    $db.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $db.TextFrame.TextRange.Text = $st.Day
    $db.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $db.TextFrame.TextRange.Font.Size = 10
    $db.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $db.TextFrame.TextRange.Font.Color.RGB = $st.Icon
    $db.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    
    $cct = $s7.Shapes.AddTextbox(1, $cx + 10, 275, 180, 165)
    $cct.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $cct.TextFrame.TextRange.Text = "$($st.Act)`r`n`r`n$($st.Desc)"
    $cct.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $cct.TextFrame.TextRange.Font.Size = 11
    $cct.TextFrame.TextRange.Font.Color.RGB = $TEXT_MUTED
    $cct.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    
    $cct.TextFrame.TextRange.Lines(1).Font.Size = 13
    $cct.TextFrame.TextRange.Lines(1).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $cct.TextFrame.TextRange.Lines(1).Font.Color.RGB = $TEXT_WHITE
}

$cf = $s7.Shapes.AddTextbox(1, 50, 470, 860, 30)
$cf.TextFrame.TextRange.Text = '* Class Goal: Invite 2-3 students to answer in English, then commit together to the 7-Day Challenge!'
$cf.TextFrame.TextRange.Font.Name = 'Segoe UI'
$cf.TextFrame.TextRange.Font.Size = 11
$cf.TextFrame.TextRange.Font.Italic = [Microsoft.Office.Core.MsoTriState]::msoTrue
$cf.TextFrame.TextRange.Font.Color.RGB = $TEXT_GOLD
$cf.TextFrame.TextRange.ParagraphFormat.Alignment = 2

Set-SpeakerNote $s7 @"
[SPEAKER 4 - CLASS INTERACTION]
Now, dear classmates, we have a question for you:
'What is one small thing you can do today to be a better friend?'
(Ask 2 or 3 classmates to raise their hands and answer in English).
Then introduce the 7-Day Challenge:
Day 1-2: Give a compliment.
Day 3-4: Offer practical help.
Day 5-6: Include someone who is sitting alone.
Day 7: Express gratitude!
"@

# -------------------------------------------------------------
# SLIDE 8: QUESTIONS AND ANSWERS (Q&A)
# -------------------------------------------------------------
$s8 = $pres.Slides.Add(8, 12)
Add-SlideBackground $s8
Add-Header $s8 'Section 07' 'Questions & Answers (Q&A)' 'Prepared responses for teachers, evaluators, and audience members'

$qas = @(
    @{
        Q = 'Q1: What is the most important quality in a friend?'
        A = 'I think trust is one of the most important qualities because friends need to feel safe and respected.'
        Detail = 'Trust is the foundation of every lasting bond. Without trust, communication becomes guarded.'
        Color = $TEXT_ACCENT
    },
    @{
        Q = 'Q2: What should you do if you disagree with a friend?'
        A = 'We should stay calm, listen to each other, and explain our opinions politely.'
        Detail = 'Disagreements are natural. The key is respecting differences instead of turning them into fights.'
        Color = $TEXT_GOLD
    },
    @{
        Q = 'Q3: How can you help a lonely student?'
        A = 'I could start a conversation, invite them to join an activity, and make them feel welcome.'
        Detail = 'Taking the first step requires courage, but it shows genuine empathy and leadership.'
        Color = $TEXT_GREEN
    },
    @{
        Q = 'Q4: Can a friendship survive a misunderstanding?'
        A = 'Yes. If both friends communicate honestly and are willing to apologise, their friendship can become stronger.'
        Detail = 'Overcoming conflicts together often makes friendships even more durable and authentic.'
        Color = $TEXT_PURPLE
    }
)

for ($i = 0; $i -lt 4; $i++) {
    $row = [Math]::Floor($i / 2)
    $col = $i % 2
    $qx = 50 + ($col * 440)
    $qy = 115 + ($row * 195)
    $qa = $qas[$i]
    
    $qc = $s8.Shapes.AddShape(5, $qx, $qy, 420, 180)
    $qc.Fill.Solid()
    $qc.Fill.ForeColor.RGB = $CARD_COLOR
    $qc.Line.ForeColor.RGB = $qa.Color
    $qc.Line.Weight = 1.2
    
    $qt = $s8.Shapes.AddTextbox(1, $qx + 15, $qy + 12, 390, 155)
    $qt.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $qt.TextFrame.TextRange.Text = "$($qa.Q)`r`n`r`nAnswer: $($qa.A)`r`n`r`n$($qa.Detail)"
    $qt.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $qt.TextFrame.TextRange.Font.Size = 10.5
    $qt.TextFrame.TextRange.Font.Color.RGB = $TEXT_MUTED
    
    $qt.TextFrame.TextRange.Lines(1).Font.Size = 12
    $qt.TextFrame.TextRange.Lines(1).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $qt.TextFrame.TextRange.Lines(1).Font.Color.RGB = $qa.Color
    
    $qt.TextFrame.TextRange.Lines(3).Font.Size = 11
    $qt.TextFrame.TextRange.Lines(3).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $qt.TextFrame.TextRange.Lines(3).Font.Color.RGB = $TEXT_WHITE
}

Set-SpeakerNote $s8 @"
[ALL MEMBERS - Q&A ROTATION]
Each group member can answer one of the questions if asked by the teacher:
- Member 1 can answer Q1 regarding trust.
- Member 2 can answer Q2 on handling disagreements calmly.
- Member 3 can answer Q3 on welcoming lonely peers.
- Member 4 can answer Q4 on repairing friendships after a misunderstanding.
"@

# -------------------------------------------------------------
# SLIDE 9: TEAM ROLES, WORK PLAN & POSTER LAYOUT
# -------------------------------------------------------------
$s9 = $pres.Slides.Add(9, 12)
Add-SlideBackground $s9
Add-Header $s9 'Section 08' 'Team Work Plan & Poster Layout' 'Organizational matrix, team agreement, and visual poster layout'

$leftBox = $s9.Shapes.AddShape(5, 50, 115, 430, 390)
$leftBox.Fill.Solid()
$leftBox.Fill.ForeColor.RGB = $CARD_COLOR
$leftBox.Line.ForeColor.RGB = $TEXT_ACCENT
$leftBox.Line.Weight = 1.2

$lbt = $s9.Shapes.AddTextbox(1, 65, 125, 400, 370)
$lbt.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
$lbt.TextFrame.TextRange.Text = "TEAM ROLES & RESPONSIBILITIES`r`n`r`n" +
"* Member 1 (Campaign Leader):`r`n  Mission, slogan, opening speech & overall coordination.`r`n`r`n" +
"* Member 2 (Rules Specialist):`r`n  Five golden rules, real-life examples & speech 2.`r`n`r`n" +
"* Member 3 (Advice & Activity):`r`n  Three tips, conditional sentences & 7-day challenge.`r`n`r`n" +
"* Member 4 (Conclusion & Q&A):`r`n  Closing call to action, Q&A moderation & poster design.`r`n`r`n" +
"Team Agreement: We listen to everyone's ideas, divide work fairly, practise together, speak English, and resolve disagreements respectfully."
$lbt.TextFrame.TextRange.Font.Name = 'Segoe UI'
$lbt.TextFrame.TextRange.Font.Size = 10.5
$lbt.TextFrame.TextRange.Font.Color.RGB = $TEXT_MUTED
$lbt.TextFrame.TextRange.Lines(1).Font.Size = 13
$lbt.TextFrame.TextRange.Lines(1).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$lbt.TextFrame.TextRange.Lines(1).Font.Color.RGB = $TEXT_ACCENT

$rightBox = $s9.Shapes.AddShape(5, 490, 115, 420, 390)
$rightBox.Fill.Solid()
$rightBox.Fill.ForeColor.RGB = RGB 24 33 51
$rightBox.Line.ForeColor.RGB = $TEXT_GOLD
$rightBox.Line.Weight = 1.5

$pm1 = $s9.Shapes.AddShape(5, 510, 130, 380, 45)
$pm1.Fill.Solid()
$pm1.Fill.ForeColor.RGB = RGB 30 58 138
$pm1.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
$pm1.TextFrame.TextRange.Text = "BE A BETTER FRIEND!`r`nBE KIND * BE HONEST * SHOW RESPECT * OFFER SUPPORT"
$pm1.TextFrame.TextRange.Font.Name = 'Segoe UI'
$pm1.TextFrame.TextRange.Font.Size = 10.5
$pm1.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pm1.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE
$pm1.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$pm2 = $s9.Shapes.AddShape(5, 510, 185, 380, 130)
$pm2.Fill.Solid()
$pm2.Fill.ForeColor.RGB = $CARD_COLOR
$pm2.Line.ForeColor.RGB = $CARD_BORDER
$pm2.TextFrame.TextRange.Text = "OUR 5 RULES`r`n1. Listen carefully.`r`n2. Be honest and trustworthy.`r`n3. Show respect.`r`n4. Offer support.`r`n5. Solve problems calmly."
$pm2.TextFrame.TextRange.Font.Name = 'Segoe UI'
$pm2.TextFrame.TextRange.Font.Size = 11
$pm2.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE
$pm2.TextFrame.TextRange.Lines(1).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pm2.TextFrame.TextRange.Lines(1).Font.Color.RGB = $TEXT_GOLD

$pm3 = $s9.Shapes.AddShape(5, 510, 325, 380, 95)
$pm3.Fill.Solid()
$pm3.Fill.ForeColor.RGB = $CARD_COLOR
$pm3.Line.ForeColor.RGB = $CARD_BORDER
$pm3.TextFrame.TextRange.Text = "3 TIPS FOR SUCCESS`r`n* Listen without judging.`r`n* Speak honestly.`r`n* Include someone who feels left out."
$pm3.TextFrame.TextRange.Font.Name = 'Segoe UI'
$pm3.TextFrame.TextRange.Font.Size = 11
$pm3.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE
$pm3.TextFrame.TextRange.Lines(1).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pm3.TextFrame.TextRange.Lines(1).Font.Color.RGB = $TEXT_GREEN

$pm4 = $s9.Shapes.AddTextbox(1, 510, 430, 380, 60)
$pm4.TextFrame.TextRange.Text = '"Be the friend you wish to have!"`r`nGrade 11 BSB-1 Project Poster Design'
$pm4.TextFrame.TextRange.Font.Name = 'Segoe UI'
$pm4.TextFrame.TextRange.Font.Size = 12
$pm4.TextFrame.TextRange.Font.Italic = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pm4.TextFrame.TextRange.Font.Color.RGB = $TEXT_GOLD
$pm4.TextFrame.TextRange.ParagraphFormat.Alignment = 2

Set-SpeakerNote $s9 @"
[SPEAKER 4 - WORK PLAN & POSTER]
Collaboration is the key to our project's success.
On the left, you can see our clear division of responsibilities and team agreement.
On the right, we showcase our classroom poster layout, designed to catch attention in school corridors and promote positive student relationships.
"@

# -------------------------------------------------------------
# SLIDE 10: FINAL CHECKLIST & CONCLUSION
# -------------------------------------------------------------
$s10 = $pres.Slides.Add(10, 12)
Add-SlideBackground $s10
Add-Header $s10 'Section 09' 'Final Checklist & Closing Message' 'Ensuring full compliance with BSB-1 criteria and final call to action'

$cbox = $s10.Shapes.AddShape(5, 50, 115, 860, 235)
$cbox.Fill.Solid()
$cbox.Fill.ForeColor.RGB = $CARD_COLOR
$cbox.Line.ForeColor.RGB = $TEXT_GREEN
$cbox.Line.Weight = 1.5

$cbt = $s10.Shapes.AddTextbox(1, 70, 125, 820, 215)
$cbt.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
$cbt.TextFrame.TextRange.Text = "BSB-1 EVALUATION CRITERIA CHECKLIST`r`n`r`n" +
"[ YES ]  The poster and presentation have a clear title and inspiring slogan.`r`n" +
"[ YES ]  All five friendship rules are clearly defined and illustrated.`r`n" +
"[ YES ]  Three actionable pieces of advice are included with grammatical examples.`r`n" +
"[ YES ]  All four team members have distinct speaking roles in fluent English.`r`n" +
"[ YES ]  Comparatives, modals, and conditionals are used correctly.`r`n" +
"[ YES ]  The group has practised pronunciation, eye contact, and timing.`r`n" +
"[ YES ]  Every team member is prepared to answer questions during Q&A."
$cbt.TextFrame.TextRange.Font.Name = 'Segoe UI'
$cbt.TextFrame.TextRange.Font.Size = 11.5
$cbt.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE

$cbt.TextFrame.TextRange.Lines(1).Font.Size = 13
$cbt.TextFrame.TextRange.Lines(1).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$cbt.TextFrame.TextRange.Lines(1).Font.Color.RGB = $TEXT_GREEN

$finalBox = $s10.Shapes.AddShape(5, 50, 365, 860, 135)
$finalBox.Fill.Solid()
$finalBox.Fill.ForeColor.RGB = RGB 24 33 51
$finalBox.Line.ForeColor.RGB = $TEXT_GOLD
$finalBox.Line.Weight = 2

$fq1 = $s10.Shapes.AddTextbox(1, 70, 375, 820, 50)
$fq1.TextFrame.TextRange.Text = '"Remember, a real friend does not have to solve every problem; sometimes, being present is enough."'
$fq1.TextFrame.TextRange.Font.Name = 'Segoe UI'
$fq1.TextFrame.TextRange.Font.Size = 15
$fq1.TextFrame.TextRange.Font.Italic = [Microsoft.Office.Core.MsoTriState]::msoTrue
$fq1.TextFrame.TextRange.Font.Color.RGB = $TEXT_WHITE
$fq1.TextFrame.TextRange.ParagraphFormat.Alignment = 2

$fq2 = $s10.Shapes.AddTextbox(1, 70, 435, 820, 50)
$fq2.TextFrame.TextRange.Text = 'THANK YOU FOR LISTENING!  *  BE KIND, BE HONEST, BE A BETTER FRIEND!'
$fq2.TextFrame.TextRange.Font.Name = 'Segoe UI'
$fq2.TextFrame.TextRange.Font.Size = 13
$fq2.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
$fq2.TextFrame.TextRange.Font.Color.RGB = $TEXT_ACCENT
$fq2.TextFrame.TextRange.ParagraphFormat.Alignment = 2

Set-SpeakerNote $s10 @"
[SPEAKER 4 - CONCLUSION]
In conclusion, we have fulfilled every requirement of the BSB-1 Project Work.
Remember our key message:
'A real friend does not have to solve every problem; sometimes, being present is enough.'
Be kind, be honest, and be a better friend!
Thank you for listening! We are now open for any questions.
"@

# Save Presentation
$outputPath = "c:\Users\asus\Desktop\Air zoon\Be_A_Better_Friend_Presentation.pptx"
if (Test-Path $outputPath) {
    Remove-Item $outputPath -Force
}
$pres.SaveAs($outputPath)

# Export slides as high-definition preview images
$previewDir = "c:\Users\asus\Desktop\Air zoon\slides_preview"
if (!(Test-Path $previewDir)) {
    New-Item -ItemType Directory -Path $previewDir | Out-Null
}
for ($s = 1; $s -le $pres.Slides.Count; $s++) {
    $imgFile = Join-Path $previewDir "slide_$s.png"
    $pres.Slides.Item($s).Export($imgFile, "PNG", 1600, 900)
}

$pres.Close()
$ppt.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[System.GC]::Collect()

Write-Host "Presentation generated successfully at: $outputPath"
Write-Host "Slide preview images saved in: $previewDir"
