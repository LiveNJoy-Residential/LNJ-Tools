# Resident Activity Audit Tool — Architecture Diagram

**Prepared for:** Partha Balakrishnan / Development Team  
**Date:** July 23, 2026  
**Project:** LNJ Resident Activity Audit Tool (Move-In & Move-Out Audit)

---

```mermaid
flowchart TD
    subgraph ResMan["ResMan — Source System"]
        R1["New and Renewed Leases Export"]
        R2["Recurring Transaction Projections"]
        R3["Rent Roll Export"]
        R4["Resident Activity Export"]
        R5["Move-Out Report — Phase 2 needed"]
        R6["Document Attachment Report — Phase 2 needed"]
    end

    subgraph Tool["Resident Activity Audit Tool — Python / Streamlit"]
        I["Data Ingestion Layer — CSV Import + Field Validation"]

        subgraph MoveIn["Move-In Audit Engine"]
            MI1["Unit and Name Match"]
            MI2["Lease Date Checks"]
            MI3["Rent and Deposit Checks"]
            MI4["Recurring Charges Check"]
            MI5["Document Presence Check — Phase 2"]
        end

        subgraph MoveOut["Move-Out Audit Engine"]
            MO1["Unit and Name Match"]
            MO2["Move-Out Date and Reason"]
            MO3["Broke Lease + Fees Check"]
            MO4["Financial Reconciliation — Balance, Collections, Refund"]
            MO5["Damage Charges vs Photos — Phase 2"]
            MO6["Document Presence Check — Phase 2"]
        end

        RPT["Report Generator — Per Property, Pass / Fail / Missing, Traceable"]
    end

    subgraph Output["Audit Report Distribution"]
        PM["Property Manager"]
        RM["Regional Manager"]
        OWN["Owner"]
        XLS["Excel + Streamlit Dashboard"]
    end

    R1 --> I
    R2 --> I
    R3 --> I
    R4 --> I
    R5 -.->|Phase 2| I
    R6 -.->|Phase 2| I
    I --> MoveIn
    I --> MoveOut
    MoveIn --> RPT
    MoveOut --> RPT
    RPT --> XLS
    XLS --> PM
    XLS --> RM
    XLS --> OWN
```

---

## Phase 1 — Automated with Existing CSV Data

| Check | Data Source |
|---|---|
| Unit number match | Leases CSV |
| Resident name match | Leases CSV |
| Lease start / end / move-in dates | Leases CSV |
| Rent charge amount | Leases CSV + Rent Roll |
| Recurring charges present | Recurring Projection CSV |
| Required deposit vs. paid | Leases CSV |
| Move-out date | Resident Activity CSV |
| Broke lease flag + fees | Leases + Transactions CSV |
| Balance / financial reconciliation | Rent Roll + Transactions CSV |

## Phase 2 — Requires New ResMan Exports

| Check | What is Needed from ResMan |
|---|---|
| Lease / ID / Insurance doc uploaded? | Document Attachment Report |
| Move-out photos present? | Document Attachment Report |
| Complete lease package uploaded? | Document Attachment Report |
