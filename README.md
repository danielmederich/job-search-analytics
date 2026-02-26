# Job Search Analytics System

A fully functional Excel workbook for managing a professional job search — with live KPI dashboards, automated scoring, skills gap analysis, and macro-powered automation.

Built using AI-assisted design: ideated in Microsoft Copilot, structured and built by Claude (Anthropic).

![Dashboard Preview](dashboard_screenshot.png)

---

## What It Does

Most people track job applications in a basic spreadsheet. This system goes further — it turns your application data into actionable analytics so you can see what's working, what isn't, and where to focus next.

**Key capabilities:**
- Live KPI dashboard showing activity, follow-up discipline, interview rate, and offer rate
- Job Search Health Score (0–100) composite metric updated in real time
- Résumé performance comparison across multiple versions
- Skills Gap Analysis engine that scores your fit for any job posting
- Automated weekly summary via macro
- Chronological timeline visualization of all job search events
- Conditional formatting that flags overdue follow-ups and upcoming interviews automatically

---

## Workbook Structure

| Sheet | Purpose |
|---|---|
| Cover | Project metadata and version info |
| Log | Primary data entry — all your applications live here |
| Dashboard | Live KPIs, Health Score, charts, and macro buttons |
| Analytics | Pivot table engine and weekly activity summary |
| Timeline | Scatter chart showing all events over time |
| Skills Gap Analysis | Weighted scoring model for job fit |
| Documentation | All formulas, named ranges, and macro code |

---

## Getting Started

### 1. Enable Macros
When you open the file, click **Enable Content** if prompted.

### 2. Install the VBA Macros

- Press `Alt + F11` to open the VBA Editor
- Click **Insert → Module**
- Copy and paste the entire code block below into the module
- Press `Ctrl + S`
- Save the file as `.xlsm` (Excel Macro-Enabled Workbook) when prompted

```vba
Option Explicit
'-----------------------------------------------------------
' RefreshDashboard — assign to [Refresh Dashboard] button
'-----------------------------------------------------------
Sub RefreshDashboard()
    On Error Resume Next
    ThisWorkbook.RefreshAll
    On Error GoTo 0
    Dim ws As Worksheet
    Dim pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    Application.CalculateFull
    MsgBox "Dashboard refreshed successfully.", vbInformation, "Refresh Complete"
End Sub
'-----------------------------------------------------------
' UpdatePivotSources — assign to [Update Pivot Tables] button
'-----------------------------------------------------------
Sub UpdatePivotSources()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim src As String
    src = "JobLogTable"
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.ChangePivotCache _
                ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=src)
        Next pt
    Next ws
    MsgBox "Pivot table sources updated.", vbInformation, "Update Complete"
End Sub
'-----------------------------------------------------------
' WeeklySummary — assign to [Weekly Summary] button
'-----------------------------------------------------------
Sub WeeklySummary()
    Dim apps As Long
    Dim interviews As Long
    Dim offers As Long
    Dim followups As Long
    apps = WorksheetFunction.CountIfs( _
        Sheets("Log").Range("A:A"), ">=" & Date - 7)
    interviews = WorksheetFunction.CountIfs( _
        Sheets("Log").Range("G:G"), "Interview Scheduled", _
        Sheets("Log").Range("A:A"), ">=" & Date - 7)
    offers = WorksheetFunction.CountIfs( _
        Sheets("Log").Range("G:G"), "Offer", _
        Sheets("Log").Range("A:A"), ">=" & Date - 7)
    followups = WorksheetFunction.CountIfs( _
        Sheets("Log").Range("H:H"), "<" & Date, _
        Sheets("Log").Range("G:G"), "Applied")
    MsgBox _
        "Weekly Summary:" & vbCrLf & vbCrLf & _
        "Applications: " & apps & vbCrLf & _
        "Interviews: " & interviews & vbCrLf & _
        "Offers: " & offers & vbCrLf & _
        "Overdue Follow-Ups: " & followups, _
        vbInformation, "Weekly Summary"
End Sub
'-----------------------------------------------------------
' ExtractSkills — reads JobDescriptionText, scans vs
' SkillsDictionary (rows 8-27), populates JobPostingSkills
' (H8:J29), rebuilds SkillsGap table (rows 32-43), and
' writes correct panel formulas for Top 5 Gaps and Strengths
' VERSION: 2026.9 — SUMPRODUCT ranking, all headers restored
'-----------------------------------------------------------
Sub ExtractSkills()
    Dim dictWeight As Object
    Dim dictCat As Object
    Dim skill As Variant
    Dim txt As String
    Dim ws As Worksheet
    Dim i As Long
    Dim rowOut As Long
    Dim gapRow As Long
    Dim cleanRow As Long
    Dim panelRow As Long

    Set ws = Sheets("Skills Gap Analysis")
    txt = ws.Range("JobDescriptionText").Value

    Set dictWeight = CreateObject("Scripting.Dictionary")
    Set dictCat = CreateObject("Scripting.Dictionary")

    ' Read SkillsDictionary — pinned to rows 8-27 only
    For i = 8 To 27
        If ws.Cells(i, 1).Value <> "" Then
            dictWeight(ws.Cells(i, 1).Value) = ws.Cells(i, 3).Value
            dictCat(ws.Cells(i, 1).Value) = ws.Cells(i, 2).Value
        End If
    Next i

    ' Clear JobPostingSkills — capped at row 29
    For i = 8 To 29
        ws.Cells(i, 8).Value = ""
        ws.Cells(i, 9).Value = ""
        ws.Cells(i, 10).Value = ""
    Next i

    ' Clear SkillsGap data rows — capped at row 43
    For i = 32 To 43
        On Error Resume Next
        ws.Cells(i, 1).Value = ""
        ws.Cells(i, 2).Value = ""
        ws.Cells(i, 3).Value = ""
        ws.Cells(i, 4).ClearContents
        ws.Cells(i, 5).ClearContents
        ws.Cells(i, 6).ClearContents
        ws.Cells(i, 7).ClearContents
        ws.Cells(i, 8).ClearContents
        On Error GoTo 0
    Next i

    ' Write matched skills to JobPostingSkills and SkillsGap simultaneously
    rowOut = 8
    gapRow = 32

    For Each skill In dictWeight.Keys
        If InStr(1, txt, skill, vbTextCompare) > 0 Then
            ' JobPostingSkills (H:J)
            ws.Cells(rowOut, 8).Value = skill
            ws.Cells(rowOut, 9).Value = dictCat(skill)
            ws.Cells(rowOut, 10).Value = dictWeight(skill)

            ' SkillsGap table (A:H)
            ws.Cells(gapRow, 1).Value = skill
            ws.Cells(gapRow, 2).Value = dictCat(skill)
            ws.Cells(gapRow, 3).Value = dictWeight(skill)
            ' Is In Resume
            ws.Cells(gapRow, 4).Formula = _
                "=IF(COUNTIF(ResumeSkills[Skill],A" & gapRow & ")>0,1,0)"
            ' Weighted Score
            ws.Cells(gapRow, 5).Formula = _
                "=C" & gapRow & "*D" & gapRow
            ' Weighted Gap
            ws.Cells(gapRow, 6).Formula = _
                "=IF(D" & gapRow & "=1,0,C" & gapRow & ")"
            ' Gap Rank — SUMPRODUCT tiebreaker, excludes empty rows
            ws.Cells(gapRow, 7).Formula = _
                "=SUMPRODUCT((SkillsGap[Weighted Gap]>F" & gapRow & ")" & _
                "*(SkillsGap[Skill]<>"""")*1)" & _
                "+SUMPRODUCT((SkillsGap[Weighted Gap]=F" & gapRow & ")" & _
                "*(SkillsGap[Skill]<>"""")" & _
                "*(SkillsGap[Skill]<A" & gapRow & "))+1"
            ' Strength Rank — SUMPRODUCT tiebreaker, excludes empty rows
            ws.Cells(gapRow, 8).Formula = _
                "=SUMPRODUCT((SkillsGap[Weighted Score]>E" & gapRow & ")" & _
                "*(SkillsGap[Skill]<>"""")*1)" & _
                "+SUMPRODUCT((SkillsGap[Weighted Score]=E" & gapRow & ")" & _
                "*(SkillsGap[Skill]<>"""")" & _
                "*(SkillsGap[Skill]<A" & gapRow & "))+1"

            rowOut = rowOut + 1
            gapRow = gapRow + 1
        End If
    Next skill

    ' Clean up leftover formula rows below matched skills — capped at row 43
    For cleanRow = gapRow To 43
        On Error Resume Next
        ws.Cells(cleanRow, 4).ClearContents
        ws.Cells(cleanRow, 5).ClearContents
        ws.Cells(cleanRow, 6).ClearContents
        ws.Cells(cleanRow, 7).ClearContents
        ws.Cells(cleanRow, 8).ClearContents
        On Error GoTo 0
    Next cleanRow

    ' Write Top 5 Skills Gaps panel formulas (rows 52-56, cols A-D)
    For panelRow = 52 To 56
        ws.Cells(panelRow, 1).Value = panelRow - 51
        ws.Cells(panelRow, 2).Formula = _
            "=IFERROR(INDEX(SkillsGap[Skill],MATCH(A" & panelRow & ",SkillsGap[Gap Rank],0)),"""")"
        ws.Cells(panelRow, 3).Formula = _
            "=IFERROR(INDEX(SkillsGap[Weight],MATCH(A" & panelRow & ",SkillsGap[Gap Rank],0)),"""")"
        ws.Cells(panelRow, 4).Formula = _
            "=IFERROR(INDEX(SkillsGap[Category],MATCH(A" & panelRow & ",SkillsGap[Gap Rank],0)),"""")"
    Next panelRow

    ' Write Top 5 Resume Strengths panel formulas (rows 60-64, cols A-D)
    For panelRow = 60 To 64
        ws.Cells(panelRow, 1).Value = panelRow - 59
        ws.Cells(panelRow, 2).Formula = _
            "=IFERROR(INDEX(SkillsGap[Skill],MATCH(A" & panelRow & ",SkillsGap[Strength Rank],0)),"""")"
        ws.Cells(panelRow, 3).Formula = _
            "=IFERROR(INDEX(SkillsGap[Weight],MATCH(A" & panelRow & ",SkillsGap[Strength Rank],0)),"""")"
        ws.Cells(panelRow, 4).Formula = _
            "=IFERROR(INDEX(SkillsGap[Category],MATCH(A" & panelRow & ",SkillsGap[Strength Rank],0)),"""")"
    Next panelRow

    ' Restore section header labels
    ws.Cells(45, 1).Value = "JOB FIT SCORE"
    ws.Cells(46, 1).Value = "Total Possible Weight:"
    ws.Cells(47, 1).Value = "Total Achieved Weight:"
    ws.Cells(48, 1).Value = "JOB FIT SCORE (0-100):"
    ws.Cells(50, 1).Value = "TOP 5 SKILLS GAPS  (Most Important Missing Skills)"
    ws.Cells(58, 1).Value = "TOP 5 RESUME STRENGTHS  (Highest Weighted Matching Skills)"

    ' Restore Job Fit Score formulas in column C
    ws.Cells(46, 3).Formula = "=SUM(JobPostingSkills[Weight])"
    ws.Cells(47, 3).Formula = "=SUM(SkillsGap[Weighted Score])"
    ws.Cells(48, 3).Formula = _
        "=IFERROR((SUM(SkillsGap[Weighted Score])/SUM(JobPostingSkills[Weight]))*100,0)"

    ' Restore panel table headers
    ws.Cells(51, 1).Value = "Rank"
    ws.Cells(51, 2).Value = "Skill"
    ws.Cells(51, 3).Value = "Weight"
    ws.Cells(51, 4).Value = "Category"

    ws.Cells(59, 1).Value = "Rank"
    ws.Cells(59, 2).Value = "Skill"
    ws.Cells(59, 3).Value = "Weight"
    ws.Cells(59, 4).Value = "Category"

    MsgBox "Skills extracted, SkillsGap rebuilt, and panels updated successfully.", _
        vbInformation, "Extraction Complete"
End Sub
```

### 3. Wire Up the Macro Buttons

The Dashboard has three buttons that need to be linked to the macros. Do this once after installing the macro code:

- On the Dashboard tab, go to **Insert → Shapes → Rounded Rectangle**
- Draw a shape over each of the three button cells (Refresh Dashboard, Weekly Summary, Update Pivot Tables)
- Type the button label inside each shape
- Right-click the shape border (not inside it) → **Assign Macro** → select the matching macro name:
  - `Refresh Dashboard` → `RefreshDashboard`
  - `Weekly Summary` → `WeeklySummary`
  - `Update Pivot Tables` → `UpdatePivotSources`

**Optional — style the buttons to match the dashboard:**
- Right-click shape → **Format Shape**
- Fill → Solid fill → `#3A4A5A`
- Line → Solid line → `#4F8A8B`
- Text → color `#C9D1D3`, Calibri, 11pt, bold

### 4. Log Your First Application
Go to the **Log** sheet and start entering your applications. Every field has dropdown validation — no free-typing required for Status, Resume Version, or Outcome.

### 5. Refresh the Dashboard
Click the **Refresh Dashboard** button on the Dashboard tab. Your KPIs, Health Score, and charts will update immediately.

---

## The Skills Gap Analysis Engine

This is the most powerful feature. For any role you're considering:

1. Go to the **Skills Gap Analysis** tab
2. Paste the full job description text into cell **A4**
3. Click **Extract Skills** (macro button) — it scans the posting against your Skills Dictionary automatically
4. Review your **Job Fit Score (0–100)** and the **Top 5 Gaps** panel
5. Use the gaps to tailor your résumé before applying

---

## Job Search Health Score

A composite 0–100 score that summarizes overall search health:

| Component | Weight |
|---|---|
| Application Activity | 30% |
| Follow-Up Discipline | 30% |
| Interview Rate | 25% |
| Offer Rate | 15% |

**Score guide:**
- 80–100 — Strong momentum
- 50–79 — Moderate, room to improve
- 0–49 — Needs attention

---

## Macro Reference

| Macro | Button | What It Does |
|---|---|---|
| `RefreshDashboard` | Refresh Dashboard | Refreshes all pivot tables and recalculates the full workbook |
| `WeeklySummary` | Weekly Summary | Pops up a summary of the last 7 days — applications, interviews, offers, overdue follow-ups |
| `UpdatePivotSources` | Update Pivot Tables | Re-links all pivot tables to the JobLogTable after adding new data |
| `ExtractSkills` | Extract Skills | Scans job description text against the Skills Dictionary and populates the Job Posting Skills table |

---

## Recommended Workflow

**Daily (5–10 min)**
- Log new applications in the Log sheet
- Set Follow-Up Dates 3–5 days out
- Run Refresh Dashboard

**Weekly (20–30 min)**
- Run Weekly Summary macro
- Review Health Score
- Check Résumé Performance chart — which version is getting interviews?
- Run Update Pivot Tables

**Monthly (30–45 min)**
- Review Skills Gap trends
- Update your résumé versions based on data
- Track Health Score improvement over time

---

## How This Was Built

This system started as a concept ideated with **Microsoft Copilot**, then the full specification was structured and handed off to **Claude (Anthropic)**, which built the complete workbook — all 7 sheets, 771 formulas, named ranges, tables, charts, conditional formatting, and macro code — in a single session.

The workflow demonstrated something genuinely useful: use AI for ideation in your existing ecosystem, then use AI for precise structured execution.

---

## Requirements

- Microsoft Excel 2019 or later (Windows recommended for full macro compatibility)
- Macros must be enabled
- No external dependencies or add-ins required

---

## Version

**2026.1** — Initial release
Designed around the principles of **Quiet Productivity** — calm visual design, low cognitive load, high clarity, audit-ready documentation.

---

## License

This project is licensed under the **MIT License**.

**What that means in plain English:** You are free to download, use, modify, and share this file for any purpose — personal or otherwise — at no cost and with no restrictions. The MIT License is one of the most open and widely used licenses in the software world. The only requirement is that the original license text is included if you redistribute it, which GitHub handles automatically.
