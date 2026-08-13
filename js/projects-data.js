/** Single source of truth — portfolio + resume projects stay in sync */
const PORTFOLIO_PROJECTS = [
  {
    title: 'Flexcube vs Oracle GL Reconciliation Automation',
    type: 'internal',
    company: 'Capri Global Capital Limited',
    description:
      'End-to-end Python automation built with vibe coding (AI coding agents in VS Code). Reconciles Flexcube Trial Balance CSV reports against Oracle Fusion GL dumps, generates formatted Excel workbooks with GL-level diffs, and sends SUCCESS/WARNING email alerts to finance teams.',
    metrics: [
      '2 legal entities — CGCL & CGHFL',
      'Daily automated GL reconciliation',
      'Flexcube closing vs Oracle net — difference tracking',
      'Auto email alerts when mismatches exceed threshold',
    ],
    workflow:
      'Ingest TB & Oracle files → Pandas grouping by GL code → cumulative Oracle net → date-shifted matching → OpenPyXL report → SMTP/Outlook notification.',
    tools: ['Python', 'Pandas', 'OpenPyXL', 'Oracle', 'VS Code', 'AI Agents'],
  },
  {
    title: 'HL RPS Cash Flow Bucketing Automation (Streamlit)',
    type: 'internal',
    company: 'Capri Global Capital Limited',
    description:
      'Interactive Streamlit application built with vibe coding in VS Code to automate Home Loan (HL) RPS cash flow bucketing. Merges multiple CSV/Excel chunk files, standardizes cashflow types, and produces aging-bucket summaries split by Principal and Interest proceeds.',
    metrics: [
      '10 aging buckets — from 0-7 days through 5+ years',
      'Multi-file merge for EMI & Amortizing loan chunks',
      'Principal + Interest pivot analysis by time bucket',
      'One-click report generation with CSV download',
    ],
    workflow:
      'Select input folder → merge CSV/XLSX chunks → clean & map cashflow types → compute extract-to-transaction day gaps → assign aging buckets → pivot principal/interest totals → export bucket summary CSV.',
    tools: ['Python', 'Pandas', 'Streamlit', 'VS Code', 'AI Agents'],
  },
  {
    title: 'Credit Card Transaction Analysis',
    type: 'public',
    link: 'https://github.com/Sant-78',
    description:
      'Comprehensive credit card monthly dashboard providing real-time insights into key performance metrics and trends, enabling stakeholders to monitor and analyze credit card operations effectively.',
    metrics: [
      '55M overall revenue',
      '7.84M total interest earned',
      '45M total transactions',
      'Blue & Silver cards — 93% of overall transactions',
    ],
    workflow:
      'Breakdown by expense type, job category, card category, and chip-usage revenue.',
    tools: ['Power BI'],
  },
];
