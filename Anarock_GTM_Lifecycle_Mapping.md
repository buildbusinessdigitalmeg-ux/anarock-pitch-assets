# Anarock GTM Stage-Gate Playbook: Developer Lifecycle Mapping

This document maps the marketing, sales, and launch velocity activities of the **Anarock GTM Stage-Gate Playbook** onto generic developer-level project lifecycles. It aligns commercial launch activities with physical construction and real estate development milestones.

---

## 1. Executive Summary & Alignment Thesis
Traditional developer-level playbooks focus on **physical design, regulatory approvals, and construction milestones** (e.g., land due diligence, shell excavation, slab casting, and occupancy certification). 

Conversely, **Anarock’s GTM Stage-Gate Playbook** is an agency-side commercial execution framework. It governs marketing efficiency, pricing architecture, broker activation, and sales velocity under an exclusive mandate. 

By mapping these activities onto a generic developer's physical timelines, we ensure that:
* **Market intelligence** feeds concept design and unit mix configurations before final building plans are frozen.
* **Pre-sales token aggregation** completes in time to satisfy cash flow thresholds required for institutional debt drawdowns.
* **Ongoing sales campaigns** are synchronized with construction collection schedules to optimize financing costs.

---

## 2. Integrated Timeline Alignment Matrix

The following table aligns the lifecycles across a unified project schedule, highlighting the overlapping stages and critical gating milestones:

| Timeframe | Developer Physical Construction | Developer Gating Milestones | Anarock Mandate GTM Playbook | Key Mandate GTM Activities & Deliverables | Decision Gates & Thresholds |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Months 1–3** | **Phase 1: Land & Feasibility**<br>*Land acquisition, due diligence & base modeling* | **Stage 1: Discovery & Screening**<br>*Zoning analysis, market feasibility screening* | **Stage 1: Pre-Mandate Discovery & Pricing Analysis**<br>*Underwrite mandate margins & local absorption* | • Competitor Price-Volume Conjoint Analysis<br>• Micro-Market Demand Heatmaps<br>• Initial Financial Hurdle Models | **Gate 1: Mandate Sign-Off**<br>• Blended Equity IRR &ge; 20%<br>• Local Overhang &le; 18 months |
| **Months 4–6** | **Phase 2: Planning & Design**<br>*Zoning approvals, master planning & utility planning* | **Stage 2: Design & Concept Approval**<br>*Massing models, architectural concepts & design reviews* | **Stage 2: GTM Strategy & Prep**<br>*Collaterals, pricing splits & agent setup* | • Barbell Pricing Model (Compact Value vs. Premium splits)<br>• Fine-tune Walk-in Genie AI prompts<br>• CP Ranker database alignment | **Gate 2: Marketing Readiness**<br>• RERA Registration active & validated<br>• Pre-Registered CPs &ge; 150<br>• Site Experience Office 100% |
| **Months 7–10** | **Phase 3: Pre-Sales Setup**<br>*Launch prep, experience office & CP kickoff* | **Stage 3: Pre-Launch Readiness**<br>*Launch checks, pricing approvals, & tenders* | **Stage 3: Pre-Launch Soft Bookings**<br>*Private reviews & early token aggregation* | • Private Concierge MVP walkthroughs<br>• Astra Platinum Lead Propensity scoring<br>• Expression of Interest (EOI) token gathering | **Gate 3: Soft Bookings Gate**<br>• Pre-sales &ge; 15% of Phase 1<br>• Tokens gathered &ge; ₹12.0 Cr |
| **Months 11–42** | **Phase 4: Structure Construction**<br>*Excavation, core structure & finishing runs* | **Stage 4: Construction & Quality QA**<br>*Superstructure execution & quality audits* | **Stage 4: Public Launch & Run**<br>*Omnichannel sales, client reporting & lead recovery* | • 360° Omnichannel campaign activation<br>• n8n Lead Revival integration (Astra Phoenix)<br>• Developer Mandate Health Balanced Scorecard | **Gate 4: Ongoing Run Gate**<br>• Weekly Sales Velocity &ge; 12/mo<br>• Overall CAC &le; 2.50% |
| **Months 42–48+** | **Phase 5: Handovers & Exit**<br>*Stabilization, key handovers & snagging* | **Stage 5: Project Handover**<br>*OC procurement, customer snagging, FM transition* | **Sustenance Phase & Handover Transition**<br>*Final handovers & mandate wrap-up* | • Final contract inventory cleanup<br>• Mandate performance dashboard audit<br>• Transition dashboard to developer CRM | **Final Exit Review**<br>• Handover NPS &ge; 70<br>• 100% receivables collected |

---

## 3. Comprehensive Project Lifecycle Flowchart

The flowchart below visualizes how the stages of the **Anarock GTM Playbook** flow horizontally alongside a generic developer's lifecycle, showing the exact dependencies and check gates:

```mermaid
flowchart TB
    %% Styles & Colors
    classDef rtStyle fill:#0f172a,stroke:#38bdf8,stroke-width:2px,color:#f8fafc;
    classDef blunexStyle fill:#1e1b4b,stroke:#818cf8,stroke-width:2px,color:#f8fafc;
    classDef anarockStyle fill:#18020c,stroke:#ff3366,stroke-width:2px,color:#f8fafc;
    classDef gateStyle fill:#1e130c,stroke:#f59e0b,stroke-width:2px,color:#f8fafc;
    classDef deliverableStyle fill:#061c15,stroke:#10b981,stroke-width:1px,color:#ecfdf5;

    %% ----------------------------------------------------
    %% SUBLAYER 1: GENERIC PHYSICAL LIFECYCLE
    %% ----------------------------------------------------
    subgraph DevConstruction ["Developer Physical Construction Lifecycle"]
        rt1["Phase 1: Land & Feasibility<br>(Months 1-3)"]:::rtStyle
        rt2["Phase 2: Master Planning<br>(Months 4-5)"]:::rtStyle
        rt3["Phase 3: RERA Approval<br>(Month 6)"]:::rtStyle
        rt4["Phase 4: Pre-Sales Setup<br>(Months 7-10)"]:::rtStyle
        rt5["Phase 5: Construction Work<br>(Months 11-42)"]:::rtStyle
        rt6["Phase 6: Handovers & Exit<br>(Months 42-48)"]:::rtStyle

        rt1 --> rt2 --> rt3 --> rt4 --> rt5 --> rt6
    end

    %% ----------------------------------------------------
    %% SUBLAYER 2: DEVELOPER STAGE-GATES
    %% ----------------------------------------------------
    subgraph DevGating ["Developer Decision Gating Framework"]
        bn1["Stage 1: Discovery & Screening<br>(Months 1-3)"]:::blunexStyle
        bn_g1{"Gate 1:<br>Screening"}:::gateStyle
        bn2["Stage 2: Design & Concept<br>(Months 4-6)"]:::blunexStyle
        bn_g2{"Gate 2:<br>Concept"}:::gateStyle
        bn3["Stage 3: Pre-Launch Readiness<br>(Months 7-10)"]:::blunexStyle
        bn_g3{"Gate 3:<br>Launch Sign-off"}:::gateStyle
        bn4["Stage 4: Construction QA<br>(Months 11-42)"]:::blunexStyle
        bn_g4{"Gate 4:<br>Pre-OC Snagging"}:::gateStyle
        bn5["Stage 5: Handover & Exit<br>(Months 42-48)"]:::blunexStyle

        bn1 --> bn_g1
        bn_g1 -->|Pass| bn2
        bn2 --> bn_g2
        bn_g2 -->|Pass| bn3
        bn3 --> bn_g3
        bn_g3 -->|Pass| bn4
        bn4 --> bn_g4
        bn_g4 -->|Pass| bn5
    end

    %% ----------------------------------------------------
    %% SUBLAYER 3: ANAROCK SALES & GTM PLAYBOOK
    %% ----------------------------------------------------
    subgraph Anarock ["Anarock GTM Stage-Gate Playbook (Exclusive Mandate)"]
        %% Stage 1
        an1["Stage 1: Pre-Mandate Discovery & Pricing Analysis<br>(Months 1-2)"]:::anarockStyle
        an1_deliv["• Price-Volume Conjoint<br>• Demand Heatmap<br>• Feasibility Models"]:::deliverableStyle
        an_g1{"Gate 1:<br>Mandate Sign-Off"}:::gateStyle

        an1 --- an1_deliv
        an1 --> an_g1

        %% Stage 2
        an2["Stage 2: GTM Strategy & Prep<br>(Months 3-4)"]:::anarockStyle
        an2_deliv["• Barbell Pricing Grid<br>• Walk-in Genie Prompts<br>• CP Ranker Database"]:::deliverableStyle
        an_g2{"Gate 2:<br>Marketing Readiness"}:::gateStyle

        an_g1 -->|Approved| an2
        an2 --- an2_deliv
        an2 --> an_g2

        %% Stage 3
        an3["Stage 3: Pre-Launch Soft Bookings<br>(Month 5)"]:::anarockStyle
        an3_deliv["• Concierge MVP Previews<br>• Astra Lead Propensity Scoring<br>• EOI Token Aggregation"]:::deliverableStyle
        an_g3{"Gate 3:<br>Soft Bookings"}:::gateStyle

        an_g2 -->|Ready| an3
        an3 --- an3_deliv
        an3 --> an_g3

        %% Stage 4
        an4["Stage 4: Public Launch Run<br>(Month 6+)"]:::anarockStyle
        an4_deliv["• 360° Omnichannel Run<br>• n8n Lead Revival (Astra Phoenix)<br>• Mandate Health Balanced Scorecard"]:::deliverableStyle
        an_g4{"Gate 4:<br>Ongoing Run Gate"}:::gateStyle

        an_g3 -->|Token Targets Met| an4
        an4 --- an4_deliv
        an4 --> an_g4
    end

    %% ----------------------------------------------------
    %% CROSS-FRAMEWORK LOGICAL INTERDEPENDENCIES
    %% ----------------------------------------------------
    %% Aligning Stage 1
    rt1 -.->|Land Feasibility| an1
    bn1 -.->|Market Research| an1
    an_g1 -.->|Informs Land Go/No-Go| bn_g1

    %% Aligning Stage 2
    rt2 -.->|Product Mix & Specs| an2
    bn2 -.->|Concept Design Layouts| an2
    an_g2 -.->|RERA approval gates| rt3

    %% Aligning Stage 3
    rt4 -.->|Broker Channel Kickoff| an3
    bn3 -.->|Tendering & Pricing card| an3
    an_g3 -.->|Validates pre-sales for| bn_g3

    %% Aligning Stage 4
    rt5 -.->|Construction collection demands| an4
    bn4 -.->|Intermediate QA Gates| an4
    an_g4 -.->|Triggers barbell flex pricing| bn_g4
```

---

## 4. Key Cross-Framework Interdependencies & Sync Points

1. **Pre-Mandate Phase (Months 1–3):**
   * **The Sync:** Strategy research team runs the *Competitor Price-Volume Conjoint Analysis* (Anarock Stage 1) to determine user willingness-to-pay (WTP) for specific smart-home configurations (Developer Stage 1).
   * **The Gating Link:** If the *Financial Hurdle Model* shows a blended IRR below 20.0%, it flags a fail for **Gate 1** of both the screening and Anarock's mandate underwriting, preventing land acquisition/JDA signing.

2. **Concept & Preparation Phase (Months 4–6):**
   * **The Sync:** While the architect files the building plans for regulatory sanctions (Developer Phase 2), Anarock designs the *Barbell Pricing Model*. By offering smaller, highly efficient compact units alongside signature penthouses, the PM locks in the sales velocity layout without degrading overall project gross realization.
   * **The Gating Link:** **Gate 2** requires a validated building plan sanction and an active RERA registration number before the 360° lead-generation engine goes live.

3. **Soft Launch & Token Aggregation (Months 7–10):**
   * **The Sync:** Anarock triggers the private *Concierge MVP Previews* to test the smart-home IoT value proposition on real buyers. Concurrently, the *Astra Platinum* engine scores lead propensity to optimize physical site-visit conversion.
   * **The Gating Link:** Proceeding to full construction and launching public sales campaigns (**Gate 3**) requires a minimum of **₹12.0 Cr** in committed token advances. This cash acts as equity cushioning to unlock the developer's construction finance drawdowns.

4. **Construction & Sustenance Run (Months 11–42):**
   * **The Sync:** As structural columns are cast floor-by-floor (Developer Phase 5 / Stage 4), Anarock manages the sales run-rate. If weekly velocity falls below 12 units or CAC exceeds 2.50%, the *Astra Phoenix* n8n webhook is triggered to automatically revive stalled leads, bypassing expensive paid acquisition.
   * **The Gating Link:** Compliance with **Gate 4** (Pre-Launch Readiness / Handover prep) requires that at least 70% of inventory is booked and collections are running above an 80% collection efficiency threshold to service active debt.
