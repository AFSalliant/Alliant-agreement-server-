const express = require(‘express’);
const cors = require(‘cors’);
const {
Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
PageBreak, UnderlineType
} = require(‘docx’);

const app = express();
app.use(cors());
app.use(express.json({ limit: ‘2mb’ }));

// ── COLORS ────────────────────────────────────────────────────────────────────
const NAVY_HEADER  = “1E4D78”;
const NAVY_TITLE   = “16355C”;
const NAVY_SECTION = “1D3659”;
const BLUE_ACCENT  = “4E80BC”;
const ROW_LIGHT    = “E9F0FA”;
const ROW_LIGHTER  = “F3F6FA”;
const ROW_MED      = “D8E4F4”;
const GRID_COLOR   = “C0D0E8”;
const PAGE_W       = 9360; // US Letter, 1” margins

// ── HELPERS ───────────────────────────────────────────────────────────────────
const cellBorder = (color) => ({
top:    { style: BorderStyle.SINGLE, size: 4, color },
bottom: { style: BorderStyle.SINGLE, size: 4, color },
left:   { style: BorderStyle.SINGLE, size: 4, color },
right:  { style: BorderStyle.SINGLE, size: 4, color },
});

const noBorder = {
top:    { style: BorderStyle.NONE, size: 0, color: “FFFFFF” },
bottom: { style: BorderStyle.NONE, size: 0, color: “FFFFFF” },
left:   { style: BorderStyle.NONE, size: 0, color: “FFFFFF” },
right:  { style: BorderStyle.NONE, size: 0, color: “FFFFFF” },
};

const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };
const LINE  = “*”.repeat(28);
const SLINE = “*”.repeat(18);
const LLINE = “_”.repeat(36);

function fmtMoney(val) {
if (!val) return “$” + LINE;
const n = parseFloat(String(val).replace(/[^0-9.]/g, “”));
if (isNaN(n)) return “$” + val;
return “$” + n.toLocaleString(“en-US”, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function fmtPct(val) {
return val ? val + “%” : “_____%”;
}

function dv(val, fallback) {
return val || (fallback !== undefined ? fallback : LINE);
}

function hRule() {
return new Paragraph({
border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY_HEADER, space: 1 } },
spacing: { before: 0, after: 100 },
children: []
});
}

function thinRule() {
return new Paragraph({
border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: GRID_COLOR, space: 1 } },
spacing: { before: 0, after: 80 },
children: []
});
}

function spacer(pts = 80) {
return new Paragraph({ spacing: { before: 0, after: pts }, children: [] });
}

function titlePara(text) {
return new Paragraph({
alignment: AlignmentType.CENTER,
spacing: { before: 60, after: 40 },
children: [new TextRun({ text, bold: true, size: 26, color: NAVY_TITLE, font: “Arial” })]
});
}

function subtitlePara(text) {
return new Paragraph({
alignment: AlignmentType.CENTER,
spacing: { before: 0, after: 30 },
children: [new TextRun({ text, size: 19, color: BLUE_ACCENT, font: “Arial” })]
});
}

function scheduleTitle(text) {
return new Paragraph({
alignment: AlignmentType.CENTER,
spacing: { before: 100, after: 60 },
children: [new TextRun({ text, bold: true, size: 21, color: NAVY_TITLE, font: “Arial” })]
});
}

function sectionHead(num, text) {
return new Paragraph({
spacing: { before: 140, after: 50 },
children: [new TextRun({ text: `${num}. ${text}`, bold: true, size: 19, color: NAVY_SECTION, font: “Arial” })]
});
}

function purchaseSaleHead(text) {
return new Paragraph({
alignment: AlignmentType.CENTER,
spacing: { before: 80, after: 50 },
children: [new TextRun({ text, bold: true, size: 19, color: NAVY_SECTION, font: “Arial”,
underline: { type: UnderlineType.SINGLE } })]
});
}

function bodyPara(runs) {
return new Paragraph({
alignment: AlignmentType.BOTH,
spacing: { before: 0, after: 60 },
children: runs.map(r => {
if (typeof r === ‘string’) return new TextRun({ text: r, size: 19, font: “Arial” });
return new TextRun({ …r, size: r.size || 19, font: “Arial” });
})
});
}

function subPara(letter, runs) {
const textRuns = typeof runs === ‘string’
? [new TextRun({ text: runs, size: 19, font: “Arial” })]
: runs.map(r => new TextRun({ …r, size: 19, font: “Arial” }));
return new Paragraph({
alignment: AlignmentType.BOTH,
spacing: { before: 0, after: 60 },
indent: { left: 400, hanging: 400 },
children: [
new TextRun({ text: `(${letter})  `, size: 19, font: “Arial” }),
…textRuns
]
});
}

function hdrRow(cols, widths) {
return new TableRow({
tableHeader: true,
children: cols.map((text, i) => new TableCell({
borders: cellBorder(NAVY_HEADER),
shading: { fill: NAVY_HEADER, type: ShadingType.CLEAR },
margins: cellMargins,
width: { size: widths[i], type: WidthType.DXA },
verticalAlign: VerticalAlign.CENTER,
children: [new Paragraph({
children: [new TextRun({ text, bold: true, color: “FFFFFF”, size: 18, font: “Arial” })]
})]
}))
});
}

function dataRow(cells, widths, shade) {
return new TableRow({
children: cells.map((cell, i) => {
const runs = Array.isArray(cell)
? cell.map(r => new TextRun({ …r, size: 18, font: “Arial” }))
: [new TextRun({ text: String(cell), size: 18, font: “Arial” })];
return new TableCell({
borders: cellBorder(GRID_COLOR),
shading: { fill: shade, type: ShadingType.CLEAR },
margins: cellMargins,
width: { size: widths[i], type: WidthType.DXA },
verticalAlign: VerticalAlign.CENTER,
children: [new Paragraph({ children: runs })]
});
})
});
}

function altRows(rows, widths) {
return rows.map((r, i) => dataRow(r, widths, i % 2 === 0 ? ROW_LIGHT : ROW_LIGHTER));
}

function sigTable(leftLabel, rightLabel, leftExtra) {
const mkLeft = () => {
const runs = [
new TextRun({ text: leftLabel, bold: true, size: 19, font: “Arial” }),
new TextRun({ text: “\n\nSignature:  “ + LLINE, size: 19, font: “Arial” }),
new TextRun({ text: “\n\nName:  “ + LLINE, size: 19, font: “Arial” }),
];
if (leftExtra) runs.push(new TextRun({ text: `\n${leftExtra}:  ${LLINE}`, size: 19, font: “Arial” }));
runs.push(new TextRun({ text: “\n\nDate:  “ + LLINE, size: 19, font: “Arial” }));
return runs;
};
const mkRight = () => [
new TextRun({ text: rightLabel, bold: true, size: 19, font: “Arial” }),
new TextRun({ text: “\n\nSignature:  “ + LLINE, size: 19, font: “Arial” }),
new TextRun({ text: “\n\nName / Title:  “ + LLINE, size: 19, font: “Arial” }),
new TextRun({ text: “\n\nDate:  “ + LLINE, size: 19, font: “Arial” }),
];
const half = Math.floor(PAGE_W / 2);
return new Table({
width: { size: PAGE_W, type: WidthType.DXA },
columnWidths: [half, PAGE_W - half],
rows: [new TableRow({ children: [
new TableCell({ borders: noBorder, width: { size: half, type: WidthType.DXA },
margins: { top: 0, bottom: 0, left: 0, right: 200 },
children: [new Paragraph({ children: mkLeft(), spacing: { after: 0 } })] }),
new TableCell({ borders: noBorder, width: { size: PAGE_W - half, type: WidthType.DXA },
margins: { top: 0, bottom: 0, left: 200, right: 0 },
children: [new Paragraph({ children: mkRight(), spacing: { after: 0 } })] }),
]})],
});
}

// ── DOCUMENT BUILDER ─────────────────────────────────────────────────────────
function buildDocument(v) {
const children = [];
const p = (…items) => children.push(…items);

const m  = (id) => fmtMoney(v[id]);
const pct = (id) => fmtPct(v[id]);
const d  = (id, fb) => dv(v[id], fb);

// Interest description
const modeDesc = v.interestMode === “simple”
? `${v.simpleRate || "___"}% per six-month Settlement Period applied to the Total Purchase Cost (not compounded)`
: `${v.monthlyRate || "___"}% per month, compounded monthly on the outstanding balance`;

// ── HEADER ──────────────────────────────────────────────────────────────────
p(titlePara(“ALLIANT LEGAL FUNDING LLC”),
subtitlePara(“Non-Recourse Purchase and Sale Agreement”),
subtitlePara(“Consumer Personal Injury Pre-Settlement Funding”),
spacer(60), hRule());

// Notice box
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: [PAGE_W],
rows: [new TableRow({ children: [new TableCell({
borders: cellBorder(NAVY_HEADER),
shading: { fill: ROW_LIGHTER, type: ShadingType.CLEAR },
margins: { top: 100, bottom: 100, left: 180, right: 180 },
width: { size: PAGE_W, type: WidthType.DXA },
children: [new Paragraph({
alignment: AlignmentType.BOTH, spacing: { before: 0, after: 0 },
children: [
new TextRun({ text: “Important Notice: “, bold: true, size: 18, font: “Arial” }),
new TextRun({ text: “This transaction is a non-recourse purchase and sale of a contingent interest in potential proceeds of a legal claim. It is not a loan. If there is no recovery from the claim, Purchaser receives nothing, except as expressly provided for fraud, material misrepresentation, or other specified default provisions set forth herein.”, size: 18, font: “Arial” }),
]
})]
})]})],
}));

p(spacer(60), thinRule(), spacer(60));

// ── OPENING ────────────────────────────────────────────────────────────────
p(titlePara(“NON-RECOURSE PURCHASE AND SALE AGREEMENT”), spacer(40));
p(bodyPara([
{ text: “THIS NON-RECOURSE PURCHASE AND SALE AGREEMENT”, bold: true },
{ text: `(the "Agreement") is made on` },
{ text: d(“agreementDate”), bold: true },
{ text: `("Date of Agreement") by and between` },
{ text: “Alliant Legal Funding LLC”, bold: true },
{ text: `, a ${d("purchaserState")} limited liability company with offices at ${d("purchaserAddress")} ("Purchaser", "us") and ` },
{ text: d(“sellerName”), bold: true },
{ text: `, residing at ${d("sellerStreet")}, ${d("sellerCityStateZip")} ("Seller", "you").` },
]));

p(spacer(40), thinRule(), spacer(40));
p(purchaseSaleHead(“Purchase and Sale”), spacer(30));
p(bodyPara([{ text: `Seller holds the Legal Claim described in Schedule 1. Pursuant to this Agreement, Seller hereby agrees to sell and assign to Purchaser, without recourse, Seller's entire right, title, and interest in the Proceeds of the Legal Claim in the amount provided for under this Agreement ("Purchased Amount"). As payment for the Purchased Amount, Purchaser agrees to pay Seller the Funded Amount set forth in the Disclosure Statement below, as further described in this Agreement.` }]));

p(spacer(40), thinRule(), spacer(60));

// ── CONSUMER DISCLOSURE TABLE ──────────────────────────────────────────────
p(scheduleTitle(“CONSUMER DISCLOSURE AND TRANSACTION SUMMARY”));
const discW = [3960, 1800, 3600];
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: discW,
rows: [
hdrRow([“Transaction Item”, “Amount”, “Notes”], discW),
…altRows([
[“Funded Amount”,                                m(“fundedAmount”),        “Gross amount advanced by Purchaser”],
[“Amount Paid Directly to Seller”,              m(“disbursalToSeller”),   “Net cash to consumer”],
[“Amount Paid to Prior Funder(s), if any”,      m(“payoffPriorFunder”),   “Leave blank if none”],
[“Origination / Application Fee”,               m(“originationFee”),      “One-time fee, if applicable”],
[“Administrative / Underwriting Fee”,           m(“adminFee”),            “One-time fee, if applicable”],
[“Annual Maintenance Fee”,                      m(“annualMaintFee”),      “Charged each year from Date of Agreement”],
[“Payment Delivery Fee”,                        m(“deliveryFee”),         “Wire / ACH / check fee, if applicable”],
[“Total Charges Disclosed at Funding”,          m(“totalCharges”),        “Includes all charges known at funding”],
[“Total Purchase Cost / Base Contracted Amount”,m(“totalPurchaseCost”),   “Funded amount plus disclosed charges”],
[“Maximum Contracted Amount (if capped)”,       “Not to exceed “ + m(“maxContractedAmount”), “If agreement includes a cap”],
[“Interest Calculation Method”,                 modeDesc,                 “Applied per Schedule 2”],
[“APR (disclosure only)”,                       pct(“apr”),               “For disclosure purposes”],
], discW),
],
}));

p(spacer(80));

// Seller acknowledgment bar
const ackW = [2400, PAGE_W - 2400];
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: ackW,
rows: [new TableRow({ children: [
new TableCell({ borders: cellBorder(GRID_COLOR), shading: { fill: ROW_MED, type: ShadingType.CLEAR },
margins: cellMargins, width: { size: ackW[0], type: WidthType.DXA },
children: [new Paragraph({ children: [new TextRun({ text: “Seller Acknowledgment”, bold: true, size: 18, font: “Arial” })] })] }),
new TableCell({ borders: cellBorder(GRID_COLOR), shading: { fill: ROW_MED, type: ShadingType.CLEAR },
margins: cellMargins, width: { size: ackW[1], type: WidthType.DXA },
children: [new Paragraph({ children: [
new TextRun({ text: “Seller Signature:  “, size: 18, font: “Arial” }),
new TextRun({ text: LLINE, size: 18, font: “Arial” }),
new TextRun({ text: “     Date:  “, size: 18, font: “Arial” }),
new TextRun({ text: SLINE, size: 18, font: “Arial” }),
]})] }),
]})]
}));

p(spacer(100));

// Payment schedule disclosure
p(new Paragraph({
spacing: { before: 100, after: 40 },
children: [new TextRun({ text: “Payment Schedule - Disclosure Summary”, bold: true, size: 19, color: NAVY_SECTION, font: “Arial” })]
}));

const psW = [2520, 3240, 3600];
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: psW,
rows: [
hdrRow([“Date Range From Funding Date”, “Contracted Amount Due if Resolved in This Interval”, “Annualized Rate / Disclosure Note”], psW),
…altRows([
[“0 \u2013 6 months”,         m(“disc_0_6”),   “For disclosure only”],
[“6 \u2013 12 months”,        m(“disc_6_12”),  “For disclosure only”],
[“12 \u2013 18 months”,       m(“disc_12_18”), “For disclosure only”],
[“18 \u2013 24 months”,       m(“disc_18_24”), “For disclosure only”],
[“24 \u2013 30 months”,       m(“disc_24_30”), “For disclosure only”],
[“30 months and thereafter”,  m(“disc_30plus”),“For disclosure only”],
], psW),
],
}));

p(spacer(80));
p(bodyPara([{ text: “This transaction is a non-recourse purchase of a contingent interest in proceeds of a legal claim. It is not a loan. If Seller complies with the Agreement and there is no recovery from the Legal Claim, Purchaser receives nothing. If the recovery is less than the Contracted Amount, Purchaser is limited to the available proceeds after payment of attorneys’ fees, court costs, and any liens or priorities that apply by law or by this Agreement.”, size: 17 }]));

p(new Paragraph({ children: [new PageBreak()] }));

// ── NUMBERED SECTIONS ──────────────────────────────────────────────────────
p(sectionHead(“1”, “Recitals”));
p(subPara(“a”, `Seller has asserted a bona fide personal injury or related civil claim described in Schedule 1 (the "Legal Claim"). Seller anticipates that the Legal Claim may result in proceeds by settlement, judgment, award, verdict, arbitration, or other resolution.`));
p(subPara(“b”, `Seller desires immediate funds for personal purposes unrelated to the prosecution of the Legal Claim, and Purchaser is willing to purchase from Seller a contingent right to receive a portion of the proceeds of the Legal Claim on the terms stated herein.`));
p(subPara(“c”, “The parties intend this transaction to be a true purchase and sale of a contingent interest in proceeds only, and not a loan, extension of credit, or assignment of the cause of action itself.”));

p(sectionHead(“2”, “Definitions”));
p(subPara(“a”, [{text: ‘“Contracted Amount”’, bold: true}, {text: “ means the amount that Purchaser is entitled to receive from the proceeds of the Legal Claim, as determined under Schedule 2 as of the date Purchaser is paid.”}]));
p(subPara(“b”, [{text: ‘“Funded Amount”’, bold: true}, {text: ` means ${m("fundedAmount")} advanced by Purchaser to or on behalf of Seller at closing, excluding disclosed charges unless expressly stated otherwise in Schedule 2.`}]));
p(subPara(“c”, [{text: ‘“Proceeds”’, bold: true}, {text: “ means all money or things of value paid or payable on account of the Legal Claim, whether by settlement, judgment, verdict, arbitration award, mediation, or otherwise; provided that Proceeds do not include attorneys’ fees or costs that are not payable to Seller.”}]));
p(subPara(“d”, [{text: ‘“Purchased Interest”’, bold: true}, {text: “ means Purchaser’s contingent right to receive the Contracted Amount from the Proceeds, subject to the non-recourse limitations of this Agreement.”}]));
p(subPara(“e”, [{text: ‘“Use Fee”’, bold: true}, {text: ` means the periodic charge calculated as follows: ${modeDesc}.`}]));
p(subPara(“f”, [{text: ‘“Settlement Period”’, bold: true}, {text: “ means each six-month interval commencing on the Date of Agreement, as further described in Schedule 2.”}]));
p(subPara(“g”, [{text: ‘“Annual Maintenance Fee”’, bold: true}, {text: ` means the recurring annual fee of ${m("annualMaintFee")}, charged each year from the Date of Agreement to defray administrative costs of maintaining Purchaser's interest.`}]));

p(sectionHead(“3”, “Nature of Transaction; Non-Recourse; No Loan”));
p(subPara(“a”, “This Agreement is a non-recourse purchase and sale. Purchaser is buying a contingent interest in Proceeds only. Purchaser is not making a loan to Seller, and Seller is not promising an unconditional repayment obligation.”));
p(subPara(“b”, “If the Legal Claim results in no Proceeds, Purchaser receives nothing and Seller owes nothing, except for damages resulting from Seller’s fraud, intentional misrepresentation, conversion of Proceeds, or other express default under this Agreement.”));
p(subPara(“c”, “If the Legal Claim results in Proceeds insufficient to pay the full Contracted Amount, Purchaser shall receive only the available Proceeds to the extent provided in this Agreement, and Seller shall have no further personal liability for any deficiency.”));
p(subPara(“d”, “Nothing in this Agreement transfers or assigns the Legal Claim itself, and Purchaser shall have no right to direct litigation strategy, settlement decisions, or attorney-client communications except as expressly authorized in writing by Seller and permitted by law.”));
p(subPara(“e”, “Seller intends this transaction to be, and agrees that this transaction is, a purchase and sale and not a loan.”));

p(sectionHead(“4”, “Purchase and Sale; Consideration; Purchaser’s Acceptance”));
p(subPara(“a”, `Subject to Purchaser's final underwriting approval and funding, Seller hereby sells, assigns, and transfers to Purchaser the Purchased Interest, and Purchaser agrees to purchase that interest, in exchange for the Funded Amount of ${m("fundedAmount")} and other consideration stated in the Consumer Disclosure and Schedule 2.`));
p(subPara(“b”, “Notwithstanding the execution of this Agreement by Seller and Purchaser, the obligations of Purchaser under this Agreement shall not be effective unless and until Purchaser has completed its review of Seller and has accepted this Agreement by delivering the Funded Amount. Prior to Purchaser’s acceptance of this Agreement, Purchaser has no obligation to make any payments or disburse any amounts to Seller and retains sole discretion over whether to do so.”));
p(subPara(“c”, “This Agreement becomes effective only when Purchaser delivers the Funded Amount pursuant to Schedule 5. Delivery of the Funded Amount will be made to Seller and/or third parties on behalf of Seller, as requested by Seller, as provided for in Schedule 5.”));
p(subPara(“d”, “Seller shall not obtain additional funding from any other source secured by the same Proceeds without Purchaser’s prior written consent, unless Purchaser has been paid in full.”));

p(sectionHead(“5”, “Use Fee and Annual Maintenance Fee”));
p(subPara(“a”, `The Use Fee shall be calculated as follows: ${modeDesc}. The Use Fee is assessed at the start of each Settlement Period beginning on the Date of Agreement.`));
p(subPara(“b”, `The Annual Maintenance Fee of ${m("annualMaintFee")} shall be added to the Purchased Amount on each anniversary of the Date of Agreement until Purchaser is paid in full.`));
p(subPara(“c”, `The Contracted Amount at any given time equals the Total Purchase Cost plus all accrued Use Fees plus all accrued Annual Maintenance Fees as of that date, subject to the Maximum Contracted Amount cap of ${m("maxContractedAmount")} stated in Schedule 2.`));

p(sectionHead(“6”, “Payment; Priority; Mechanics of Remittance”));
p(subPara(“a”, “Immediately upon receipt of any Proceeds, Seller shall pay, or cause Seller’s Attorney to pay, Purchaser the Contracted Amount from the first Proceeds available after deduction of attorneys’ fees, court costs, and any amounts having priority by mandatory operation of law or by express agreement acknowledged in writing by Purchaser.”));
p(subPara(“b”, “Seller shall not be entitled to receive any portion of the applicable Proceeds unless and until Purchaser has been paid the amount then due under Schedule 2, except to the extent a lesser amount is required because available Proceeds are insufficient.”));
p(subPara(“c”, “If any Proceeds are delivered directly to Seller, Seller shall hold them in trust for Purchaser to the extent of Purchaser’s Purchased Interest and shall deliver the same to Purchaser within five (5) business days.”));
p(subPara(“d”, “Seller retains the option to extinguish Seller’s obligation under this Agreement at any time before resolution of the Legal Claim by paying the then-current Contracted Amount shown in Schedule 2 for the applicable interval, without any prepayment penalty beyond the applicable scheduled amount.”));
p(subPara(“f”, “If Seller decides to proceed without an attorney, Seller agrees to: (1) authorize Purchaser to put any third parties on notice of Purchaser’s interest; (2) irrevocably instruct any such third parties to directly remit any Proceeds to Purchaser before remitting any amounts to Seller; and (3) hold such third parties harmless in regard to remitting Proceeds to Purchaser in satisfaction of Seller’s obligation under this Agreement.”));

p(new Paragraph({ children: [new PageBreak()] }));

p(sectionHead(“7”, “Seller Representations and Warranties”));
p(bodyPara([{ text: “Seller represents and warrants the following:” }]));
p(subPara(“a”, “Seller is fully conversant with the English language and all oral discussions with regard to this Agreement were conducted in English, or Seller requested that discussions be provided in another language, and they were conducted accordingly.”));
p(subPara(“b”, “Seller is the lawful owner of the Legal Claim and has full authority and capacity to enter into this Agreement.”));
p(subPara(“c”, “The Legal Claim was asserted in good faith, and no portion of the Funded Amount will be used to pay attorneys’ fees, experts, filing fees, or other litigation expenses intended to support or maintain the Legal Claim.”));
p(subPara(“d”, “All information provided by Seller to Purchaser is true, complete, and not materially misleading; Seller has disclosed all prior funding, liens, letters of protection, bankruptcy proceedings, and other matters that could materially affect the value or collectability of Proceeds.”));
p(subPara(“e”, “Except as disclosed in writing, Seller has not sold, assigned, pledged, transferred, or encumbered the Legal Claim or any Proceeds other than contingency fees and case costs payable to Seller’s Attorney and statutory or court-recognized liens.”));
p(subPara(“f”, “There are no outstanding federal, state, or local tax liens against Seller, and there are no lawsuits pending or threatened against Seller that would materially affect the Proceeds.”));
p(subPara(“g”, “Seller is not indebted to any present or former spouse for support, maintenance, or similar obligations in any amount that would constitute a lien or claim against the Proceeds.”));
p(subPara(“h”, “Seller has been advised by Purchaser to discuss this matter with and to review this Agreement with an attorney prior to signing, and Seller has either received such counsel or expressly waived it.”));

p(sectionHead(“8”, “Seller Covenants”));
p(subPara(“a”, “Seller shall cooperate in good faith with the continued prosecution and resolution of the Legal Claim and shall not intentionally impair, compromise, conceal, divert, or delay Proceeds subject to Purchaser’s Purchased Interest.”));
p(subPara(“b”, “Seller shall promptly notify Purchaser of any change in attorney, contact information, settlement, dismissal, appeal, bankruptcy filing, or material development affecting the Legal Claim.”));
p(subPara(“c”, “Seller shall not obtain additional case funding secured by the same Proceeds without Purchaser’s prior written consent unless Purchaser has been paid in full or expressly subordinated in writing.”));
p(subPara(“d”, “Seller will provide any new, substitute, or additional attorney representing Seller in the Legal Claim with a signed Irrevocable Letter of Direction and will promptly notify Purchaser of any change in counsel.”));

p(sectionHead(“9”, “No Control Over Claim; No Legal Services”));
p(subPara(“a”, “Purchaser is not a law firm, is not giving legal advice, and is not assuming any duty to prosecute, settle, or manage the Legal Claim.”));
p(subPara(“b”, “All litigation and settlement decisions remain solely with Seller and Seller’s Attorney. Purchaser shall not attempt to direct or control those decisions.”));

p(sectionHead(“10”, “Legal Claim Information Authorization”));
p(subPara(“a”, “Seller authorizes Purchaser to obtain information concerning the facts upon which the Legal Claim is based, including copies of pleadings, motions, rulings, and all information filed or issued in connection with the Legal Claim, as well as periodic reports on the procedural progress of the Legal Claim, to the extent permitted by law.”));
p(subPara(“b”, “Seller further authorizes Purchaser to obtain from any third party, including any applicable insurance carrier, defense counsel, or court, any non-privileged information Purchaser reasonably deems necessary to monitor or protect its Purchased Interest in the Proceeds.”));

p(sectionHead(“11”, “Security Interest; UCC; Protective Measures”));
p(subPara(“a”, “As additional protection for Purchaser’s Purchased Interest, and only to the extent permitted by applicable law, Seller grants Purchaser a security interest in the Proceeds of the Legal Claim.”));
p(subPara(“b”, “Seller authorizes Purchaser to file UCC financing statements, amendments, notices of assignment, or similar protective filings describing Purchaser’s interest in the Proceeds, but not in the cause of action itself.”));

p(sectionHead(“12”, “Bankruptcy; Death; Successors”));
p(subPara(“a”, “If Seller becomes a debtor in any bankruptcy or insolvency proceeding before Purchaser has been paid, Seller shall cause the Purchased Interest in the Proceeds to be described as an asset of Purchaser in any schedule or document filed in connection with such proceeding, and not as a debt or obligation of Seller.”));
p(subPara(“b”, “If Seller dies before the Legal Claim is resolved, this Agreement shall bind Seller’s estate, heirs, personal representatives, and successors. Purchaser shall receive the Contracted Amount prior to any distributions to beneficiaries from the Proceeds; however, the estate’s obligation is limited solely to the Proceeds and not as a general recourse claim against other estate assets.”));

p(sectionHead(“13”, “Binding Effect; Assignment; Successor Purchaser”));
p(subPara(“a”, “This Agreement shall be binding upon Seller and upon each successor or assignee of Seller, and shall be binding upon Purchaser and each successor and assignee of Purchaser. Seller may not assign this Agreement or any of Seller’s rights or obligations hereunder.”));
p(subPara(“b”, “Purchaser may assign, sell, or pledge this Agreement and all of Purchaser’s rights hereunder to another person or entity without Seller’s consent.”));
p(subPara(“c”, `Seller acknowledges and agrees that all of Purchaser's right, title, and interest under this Agreement may be assigned by Purchaser to a third party ("Successor Purchaser") and that from and after receipt of notice of any such assignment, all instructions, notices, waivers, or demands hereunder shall be effective only if signed in writing by Successor Purchaser, and Seller and Seller's Attorney shall cause all payments due under this Agreement to be paid solely to Successor Purchaser.`));

p(new Paragraph({ children: [new PageBreak()] }));

p(sectionHead(“14”, “Events of Default; Remedies; Breach”));
p(subPara(“a”, “An Event of Default occurs if Seller intentionally breaches a material covenant of this Agreement; commits fraud or material misrepresentation; diverts, conceals, or converts Proceeds; or receives Proceeds and fails to remit Purchaser’s share within the time required after written demand.”));
p(subPara(“b”, “Seller shall be fully and personally liable to Purchaser for the Contracted Amount in the event that Seller makes any false, misleading, or untrue representations, purposefully takes actions to circumvent Seller’s obligations, or breaches any of the terms of this Agreement.”));
p(subPara(“c”, “In the event that misrepresentations or fraudulent actions by Seller prevent Purchaser from receiving its interest in the Proceeds, Seller shall indemnify and hold Purchaser harmless from and against any losses, costs, and expenses (including legal fees) incurred by Purchaser.”));

p(sectionHead(“15”, “Right of Cancellation”));
p(subPara(“a”, `Seller may cancel this Agreement without penalty or further obligation by delivering written notice of cancellation and returning the full Funded Amount of ${m("fundedAmount")} to Purchaser within five (5) business days after Seller receives funds, in immediately available funds, certified funds, or by return of Purchaser's uncashed check, as applicable.`));
p(subPara(“b”, “To be effective, cancellation must be made by: (i) delivering Purchaser’s uncashed check in person to Purchaser’s offices within five (5) business days of disbursement; or (ii) mailing a notice of cancellation together with return of the full Funded Amount by insured, registered, or certified U.S. mail, postmarked within five (5) business days of receiving funds.”));

p(sectionHead(“16”, “Attorney’s Fees”));
p(subPara(“a”, “All costs and expenses, filing fees, and legal fees of Purchaser incurred under this Agreement shall be the sole responsibility of Purchaser. Seller shall be solely responsible for the payment of Seller’s own legal fees.”));

p(sectionHead(“17”, “Governing Law”));
p(subPara(“a”, `This Agreement shall be governed by the law of ${d("governingLawState")}, without regard to conflict-of-laws principles. The Arbitration Clause in this Agreement is governed by the Federal Arbitration Act ("FAA") to the extent applicable.`));
p(subPara(“b”, `To the extent any claim is found not subject to the Arbitration Clause, such claim shall be instituted in a court of competent jurisdiction in ${d("venueDisputeRes")}, and Seller agrees to subject himself or herself to the jurisdiction of that court.`));

p(sectionHead(“18”, “Arbitration Clause”), spacer(40));

const arbRows = [];
const arbHeader = new TableRow({ children: [new TableCell({
borders: { top: { style: BorderStyle.SINGLE, size: 8, color: NAVY_HEADER }, bottom: { style: BorderStyle.NONE, size: 0, color: “FFFFFF” }, left: { style: BorderStyle.SINGLE, size: 8, color: NAVY_HEADER }, right: { style: BorderStyle.SINGLE, size: 8, color: NAVY_HEADER } },
shading: { fill: ROW_MED, type: ShadingType.CLEAR },
margins: { top: 100, bottom: 100, left: 180, right: 180 },
width: { size: PAGE_W, type: WidthType.DXA },
children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: “ARBITRATION CLAUSE”, bold: true, size: 20, font: “Arial” })] })]
})]});
arbRows.push(arbHeader);

const arbBodyItems = [
[{ text: “Read this Arbitration Clause carefully. It significantly affects your rights in any dispute with us.”, bold: true }],
[{ text: “* EITHER YOU OR WE MAY CHOOSE TO HAVE ANY DISPUTE BETWEEN US DECIDED BY ARBITRATION AND NOT IN COURT.”, bold: true }],
[{ text: “* IF A DISPUTE IS ARBITRATED, YOU AND WE WILL EACH GIVE UP OUR RIGHT TO A TRIAL BY THE COURT OR A JURY TRIAL.”, bold: true }],
[{ text: “* YOU WILL GIVE UP YOUR RIGHT TO PARTICIPATE AS A CLASS REPRESENTATIVE OR CLASS MEMBER ON ANY CLASS CLAIM YOU MAY HAVE AGAINST US.”, bold: true }],
[{ text: “Governing Law. “, bold: true }, { text: `This Arbitration Clause and any arbitration conducted hereunder is governed by the Federal Arbitration Act ("FAA").` }],
[{ text: “Covered Disputes. “, bold: true }, { text: `The term "dispute" includes any claim or dispute, in contract, tort, or otherwise, between Seller and Purchaser arising from or relating in any manner to this Agreement or any resulting relationship.` }],
[{ text: “Arbitration Procedures. “, bold: true }, { text: `Any dispute will, at either party's election, be resolved by neutral, binding arbitration and not by court action. Any dispute will be arbitrated on an individual basis and not as a class action. Seller expressly waives any right to arbitrate a class action (the "Class Action Waiver"). Seller may choose the applicable rules of the American Arbitration Association (AAA, www.adr.org) or JAMS (www.jamsadr.com). The arbitrator must be an attorney or retired judge and must issue a written award.` }],
[{ text: “Fees. “, bold: true }, { text: “If Seller demands arbitration first, Seller will pay the filing fee if required. Purchaser will advance and/or pay any other fees required by the organization’s rules. The arbitrator’s award is final and binding.” }],
[{ text: “Small Claims. “, bold: true }, { text: “Purchaser waives the right to require Seller to arbitrate an individual claim if the amount qualifies as a small claim under applicable law.” }],
[{ text: “Right to Opt Out. “, bold: true }, { text: “Seller may opt out of this Arbitration Clause by sending written notice to Purchaser’s address on the first page of this Agreement by mail, received by Purchaser no later than 45 days after the Date of Agreement.” }],
];

arbBodyItems.forEach((runs, i) => {
const isLast = i === arbBodyItems.length - 1;
arbRows.push(new TableRow({ children: [new TableCell({
borders: {
top:    { style: BorderStyle.NONE, size: 0, color: “FFFFFF” },
bottom: isLast ? { style: BorderStyle.SINGLE, size: 8, color: NAVY_HEADER } : { style: BorderStyle.NONE, size: 0, color: “FFFFFF” },
left:   { style: BorderStyle.SINGLE, size: 8, color: NAVY_HEADER },
right:  { style: BorderStyle.SINGLE, size: 8, color: NAVY_HEADER },
},
shading: { fill: i % 2 === 0 ? ROW_LIGHTER : ROW_LIGHT, type: ShadingType.CLEAR },
margins: { top: 60, bottom: 60, left: 180, right: 180 },
width: { size: PAGE_W, type: WidthType.DXA },
children: [new Paragraph({
alignment: AlignmentType.BOTH, spacing: { before: 0, after: 0 },
children: runs.map(r => new TextRun({ …r, size: 18, font: “Arial” }))
})]
})]}));
});

p(new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: [PAGE_W], rows: arbRows }));
p(spacer(80));

p(sectionHead(“19”, “Miscellaneous”));
p(subPara(“a”, “This Agreement, together with all schedules and signed acknowledgments, constitutes the entire agreement between the parties concerning the subject matter and supersedes all prior oral or written discussions, understandings, negotiations, or agreements.”));
p(subPara(“b”, “No waiver is effective unless in writing. No failure or delay in enforcement constitutes a waiver.”));
p(subPara(“c”, “If any provision is held invalid or unenforceable, the remaining provisions shall remain in effect to the fullest extent permitted by law.”));
p(subPara(“d”, “This Agreement may be executed in counterparts and by electronic signature, each of which is deemed an original and together constitute one instrument.”));
p(subPara(“e”, “This Agreement may not be modified or terminated orally, but only by a written agreement duly executed by Seller and Purchaser.”));

p(spacer(80), hRule(), spacer(60));
p(bodyPara([
{ text: “DO NOT SIGN THIS AGREEMENT IF IT CONTAINS BLANK SPACES. “, bold: true },
{ text: “Seller acknowledges the right to review this Agreement with counsel and to receive a fully completed copy at the time of signing. “ },
{ text: “IN WITNESS WHEREOF”, bold: true },
{ text: “, Seller hereby executes this Agreement.” },
]));
p(spacer(120));
p(sigTable(“SELLER / CLAIMANT”, “PURCHASER \u2014 ALLIANT LEGAL FUNDING LLC”));

p(new Paragraph({ children: [new PageBreak()] }));

// ── SCHEDULE 1 ─────────────────────────────────────────────────────────────
p(scheduleTitle(“SCHEDULE 1”), scheduleTitle(“Claim Information”), hRule(), spacer(60));
const s1W = [3240, PAGE_W - 3240];
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: s1W,
rows: altRows([
[[{ text: “Seller Name”, bold: true }],                   d(“sellerName”)],
[[{ text: “Date of Loss / Accident”, bold: true }],       d(“dateOfLoss”)],
[[{ text: “Court / Venue”, bold: true }],                 d(“courtVenue”)],
[[{ text: “Index / Case Number”, bold: true }],           d(“indexCaseNumber”)],
[[{ text: “Caption”, bold: true }],                       d(“claimCaption”)],
[[{ text: “Defendant(s) / Carrier(s)”, bold: true }],     d(“defendants”)],
[[{ text: “Adjuster or Claim Number”, bold: true }],      d(“adjusterClaimNum”)],
[[{ text: “Seller’s Attorney / Firm”, bold: true }],      `${d("attorneyName")}, ${d("attorneyFirm")}`],
[[{ text: “Known Liens / Prior Funding / Notes”, bold: true }], d(“knownLiens”, “None”)],
], s1W),
}));

p(new Paragraph({ children: [new PageBreak()] }));

// ── SCHEDULE 2 ─────────────────────────────────────────────────────────────
p(scheduleTitle(“SCHEDULE 2”), scheduleTitle(“Payment Schedule and Pricing Table”), hRule(), spacer(60));
p(bodyPara([{ text: `Interest Calculation Method: ${modeDesc}` }]));
p(spacer(40));
const s2W = [2520, 2880, 3960];
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: s2W,
rows: [
hdrRow([“Interval”, “Contracted Amount Due”, “Method / Note”], s2W),
…altRows([
[“0 \u2013 6 months”,   m(“s2_0_6”),       “Per interest method above”],
[“6 \u2013 12 months”,  m(“s2_6_12”),      “Per interest method above”],
[“12 \u2013 18 months”, m(“s2_12_18”),     “Per interest method above”],
[“18 \u2013 24 months”, m(“s2_18_24”),     “Per interest method above”],
[“24 \u2013 30 months”, m(“s2_24_30”),     “Per interest method above”],
[“30 \u2013 36 months”, m(“s2_30_36”),     “Per interest method above”],
[“Thereafter”,          m(“s2_thereafter”),“Not to exceed cap”],
], s2W),
],
}));
p(spacer(80));
const s2eW = [3960, PAGE_W - 3960];
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: s2eW,
rows: altRows([
[[{ text: “APR (disclosure only)”, bold: true }],        pct(“apr”)],
[[{ text: “Annual Maintenance Fee”, bold: true }],       `${m("annualMaintFee")} per year`],
[[{ text: “Maximum Contracted Amount Cap”, bold: true }],m(“maxContractedAmount”)],
[[{ text: “Governing Law”, bold: true }],                `State of ${d("governingLawState")}`],
[[{ text: “Venue / Dispute Resolution”, bold: true }],   d(“venueDisputeRes”)],
], s2eW),
}));

p(new Paragraph({ children: [new PageBreak()] }));

// ── SCHEDULE 3 ─────────────────────────────────────────────────────────────
p(scheduleTitle(“SCHEDULE 3”), scheduleTitle(“Irrevocable Letter of Direction to Counsel”), hRule(), spacer(60));
p(bodyPara([{ text: `To:  ${d("attorneyName")},  Esq. / Firm:  ${d("attorneyFirm")}` }]));
p(bodyPara([{ text: d(“attorneyAddress”) }]));
p(spacer(40));
const s3Paras = [
`I, ${d("sellerName")} ("Seller"), irrevocably direct you, and any substitute, successor, or additional counsel representing me in the Legal Claim described in Schedule 1, to protect and satisfy the Purchased Interest of Alliant Legal Funding LLC under the Non-Recourse Purchase and Sale Agreement dated ${d("agreementDate")} ("Agreement").`,
“You are hereby authorized and directed, to the extent permitted by law and your professional duties, to cooperate with Purchaser by providing, upon request, any non-privileged information relating to the Legal Claim, including status updates and the gross settlement amount for internal purposes only.”,
“Upon resolution of the Legal Claim, you will notify Purchaser of such resolution, request a payoff letter, and confirm the Contracted Amount before final distribution.”,
“From the first Proceeds available after attorneys’ fees, court costs, and amounts having priority by law, you are instructed to remit to Purchaser the Contracted Amount then due under Schedule 2 before distributing remaining funds to me, unless Purchaser has confirmed a lesser amount in writing. If there is a dispute as to the amount due, you will hold the full amount of Proceeds in trust for Purchaser pending resolution, except as required by law.”,
“Before permitting me to take any further funding of a like or similar nature from any source, Purchaser must be paid in full or give written permission for the additional funding.”,
“You will promptly notify Purchaser if you are no longer representing me (within 48 hours of withdrawal), or if any material development affects the Legal Claim.”,
“This direction is coupled with my obligations under the Agreement and shall remain in effect, and shall be binding on any subsequent attorney retained to represent me in the Legal Claim, until Purchaser confirms in writing that it has been paid in full or otherwise releases its interest.”,
];
for (const t of s3Paras) { p(bodyPara([{ text: t }])); p(spacer(30)); }
p(spacer(80), sigTable(“SELLER”, “ACKNOWLEDGED BY SELLER’S ATTORNEY”));

p(new Paragraph({ children: [new PageBreak()] }));

// ── SCHEDULE 4 ─────────────────────────────────────────────────────────────
p(scheduleTitle(“SCHEDULE 4”), scheduleTitle(“Attorney Acknowledgment and Undertaking”), hRule(), spacer(60));
p(bodyPara([{ text: “The undersigned attorney or law firm acknowledges receipt of the Agreement and agrees, to the extent permitted by applicable law and rules of professional conduct, as follows:” }]));
p(spacer(40));
const attyItems = [
`I represent Seller in the Legal Claim identified in Schedule 1, or I am authorized to sign on behalf of the law firm that does.`,
“My fee arrangement is contingent, and I will disburse any Proceeds through my attorney trust account in the ordinary course of representation.”,
“I shall not be paid or offered commissions or referral fees related to this Agreement, and I do not have a financial interest in Purchaser.”,
“I am following the instructions given to me by Seller in the Irrevocable Letter of Direction with regard to the Agreement.”,
“I will not disburse Proceeds payable to Seller until Purchaser’s Contracted Amount has been satisfied in accordance with the Agreement, subject to attorneys’ fees, court costs, and any amounts having priority by law.”,
“I will notify Purchaser of any settlement, award, verdict, substitution of counsel, or other material resolution event and will immediately contact Purchaser, but no later than 10 days from the date of resolution, to request a current payoff amount before final distribution.”,
“I have not received and will not accept direction from Purchaser concerning litigation strategy or settlement decisions, and I understand that Purchaser is not counsel to Seller.”,
`I understand that marking a check or accompanying letter "in full satisfaction" will not have legal effect absent written confirmation from Purchaser, and that Purchaser is authorized to deposit such check without prejudice to its rights to collect in full.`,
“I will not participate in or acknowledge any future funding or sale of any potential interest in the Proceeds of the Legal Claim except with the prior written approval of Purchaser and Seller.”,
“If for any reason I no longer represent Seller in connection with the Legal Claim, I will promptly notify Purchaser and provide any insurance, new attorney, or other information reasonably requested by Purchaser to allow Purchaser to protect its interest.”,
“I acknowledge that Purchaser has relied upon this Acknowledgment in entering into the Agreement and providing funding to Seller.”,
“This acknowledgment is limited to the handling of proceeds and does not create personal liability of counsel beyond duties voluntarily undertaken herein and under applicable law.”,
];
for (let i = 0; i < attyItems.length; i++) {
p(subPara(i + 1, attyItems[i]));
}
p(spacer(80));
p(bodyPara([{ text: “Known liens / letters of protection / prior funding (exclusive of attorneys’ fees and costs):” }]));
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: [PAGE_W],
rows: [0,1,2].map(() => new TableRow({ children: [new TableCell({
borders: cellBorder(GRID_COLOR),
shading: { fill: ROW_LIGHTER, type: ShadingType.CLEAR },
width: { size: PAGE_W, type: WidthType.DXA },
margins: { top: 120, bottom: 120, left: 120, right: 120 },
children: [new Paragraph({ children: [new TextRun({ text: “ “, size: 18 })] })]
})]}))
}));
p(spacer(80));
p(bodyPara([{ text: “Method to follow up on Legal Claim:” }]));
const ctW = [720, 3060, 720, 3060];
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: ctW,
rows: [new TableRow({ children: [
new TableCell({ borders: noBorder, width: { size: ctW[0], type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 0, right: 60 },
children: [new Paragraph({ children: [new TextRun({ text: “Email:”, bold: true, size: 18, font: “Arial” })] })] }),
new TableCell({ borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.SINGLE, size: 4, color: GRID_COLOR }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
width: { size: ctW[1], type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 0, right: 180 },
children: [new Paragraph({ children: [new TextRun({ text: d(“attorneyEmail”), size: 18, font: “Arial” })] })] }),
new TableCell({ borders: noBorder, width: { size: ctW[2], type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 0, right: 60 },
children: [new Paragraph({ children: [new TextRun({ text: “Tel:”, bold: true, size: 18, font: “Arial” })] })] }),
new TableCell({ borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.SINGLE, size: 4, color: GRID_COLOR }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
width: { size: ctW[3], type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 0, right: 0 },
children: [new Paragraph({ children: [new TextRun({ text: d(“attorneyTel”), size: 18, font: “Arial” })] })] }),
]})]
}));
p(spacer(120), sigTable(“SELLER’S ATTORNEY / LAW FIRM”, “PURCHASER ACKNOWLEDGMENT”, “Name / Firm”));

p(new Paragraph({ children: [new PageBreak()] }));

// ── SCHEDULE 5 ─────────────────────────────────────────────────────────────
p(scheduleTitle(“SCHEDULE 5”), scheduleTitle(“Funding Disbursement \u2014 Fee Schedule and Payment Instructions”), hRule(), spacer(60));
p(new Paragraph({
spacing: { before: 0, after: 60 },
children: [new TextRun({ text: “Disbursement Fee Schedule”, bold: true, size: 19, color: NAVY_SECTION, font: “Arial” })]
}));
const feeW = [2160, 900, 1800, 4500];
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: feeW,
rows: [
hdrRow([“Payment Method”, “Fee”, “Estimated Receipt”, “Notes”], feeW),
…altRows([
[“ACH / Wire Transfer”,    “$30.00”, “Next business day”,        “Purchaser is not responsible for holds or additional fees charged by the receiving bank.”],
[“Check \u2014 Express”,   “$15.00”, “2 business days”,          “No FedEx delivery to P.O. Boxes.”],
[“Check \u2014 Overnight”, “$30.00”, “Next business day”,        “No FedEx delivery to P.O. Boxes.”],
[“Check \u2014 Saturday”,  “$45.00”, “Saturday”,                 “No FedEx delivery to P.O. Boxes.”],
[“Check \u2014 USPS”,      “Free”,   “5\u20137 business days”,   “No guarantee of postal timing after mailing.”],
[“Zelle (if offered)”,     “None”,   “Same day or next business day”, “Transfers are final once sent.”],
], feeW),
],
}));
p(spacer(100));
p(new Paragraph({
spacing: { before: 0, after: 40 },
children: [new TextRun({ text: “Payment Method Selection”, bold: true, size: 19, color: NAVY_SECTION, font: “Arial” })]
}));
p(bodyPara([{ text: “Complete only one section below. Failure to complete this page completely and accurately may result in a delay of funding.” }]));
p(spacer(60));
p(bodyPara([{ text: “\u2610  ACH / Wire Transfer”, bold: true }]));
const achW = [3600, PAGE_W - 3600];
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: achW,
rows: altRows([
[[{ text: “Name of Account Holder”, bold: true }],                   v.achAccountHolder || LINE],
[[{ text: “Bank Name”, bold: true }],                                v.achBankName || LINE],
[[{ text: “Bank ABA / Routing Number (9 digits only)”, bold: true }],v.achRoutingNumber || LINE],
[[{ text: “Bank Account Number”, bold: true }],                      v.achAccountNumber || LINE],
[[{ text: “Account Type”, bold: true }],                             v.achAccountType ? `\u2611  ${v.achAccountType}` : “\u2610  Checking     \u2610  Savings”],
], achW),
}));
p(spacer(80));
p(bodyPara([{ text: “\u2610  Paper Check”, bold: true }]));
const chkW = [3600, PAGE_W - 3600];
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: chkW,
rows: altRows([
[[{ text: “Issue Check To”, bold: true }],       v.checkPayee || LINE],
[[{ text: “Street Address”, bold: true }],       v.checkStreet || LINE],
[[{ text: “City, State, Zip Code”, bold: true }],v.checkCityStateZip || LINE],
[[{ text: “Delivery Method”, bold: true }],      v.checkDelivery || “\u2610  Express   \u2610  Overnight   \u2610  Saturday   \u2610  USPS”],
], chkW),
}));
p(spacer(100));
p(new Paragraph({
spacing: { before: 0, after: 40 },
children: [new TextRun({ text: “Seller Personal Information”, bold: true, size: 19, color: NAVY_SECTION, font: “Arial” })]
}));
const piW = [3240, PAGE_W - 3240];
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: piW,
rows: altRows([
[[{ text: “Social Security Number”, bold: true }], v.sellerSSN || LINE],
[[{ text: “Mailing Address”, bold: true }],        v.sellerMailingAddr || LINE],
[[{ text: “City, State, Zip Code”, bold: true }],  v.sellerMailCSZ || LINE],
[[{ text: “Telephone Number”, bold: true }],       v.sellerPhone || LINE],
[[{ text: “Email Address”, bold: true }],          v.sellerEmail || LINE],
], piW),
}));
p(spacer(60));
p(bodyPara([{ text: “Please attach a copy of your driver’s license or government-issued state ID to this Agreement.”, bold: true }]));
p(spacer(60));
p(bodyPara([{ text: `By signing below, Seller authorizes Alliant Legal Funding LLC and its successors and assigns to disburse the Funded Amount of ${m("fundedAmount")} through the selected method above. Seller understands that if the information provided is incorrect or incomplete, Seller will be responsible for any additional charges incurred as a result.`, size: 17 }]));
p(spacer(100));

const finW = Math.floor(PAGE_W / 2);
p(new Table({
width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: [finW, PAGE_W - finW],
rows: [new TableRow({ children: [
new TableCell({ borders: noBorder, width: { size: finW, type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 0, right: 200 },
children: [new Paragraph({ children: [
new TextRun({ text: “Seller Signature:”, bold: true, size: 19, font: “Arial” }),
new TextRun({ text: `\n\n${LLINE}\n\nDate:  ${LLINE}`, size: 19, font: “Arial” }),
], spacing: { after: 0 } })] }),
new TableCell({ borders: noBorder, width: { size: PAGE_W - finW, type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 200, right: 0 },
children: [new Paragraph({ children: [
new TextRun({ text: “Purchaser Approval:”, bold: true, size: 19, font: “Arial” }),
new TextRun({ text: `\n\n${LLINE}\n\nFunding Date:  ${LLINE}`, size: 19, font: “Arial” }),
], spacing: { after: 0 } })] }),
]})]
}));

// ── BUILD DOCUMENT ─────────────────────────────────────────────────────────
return new Document({
styles: {
default: {
document: { run: { font: “Arial”, size: 19 } },
paragraph: { spacing: { line: 276, lineRule: “auto” } }
},
},
sections: [{
properties: {
page: {
size: { width: 12240, height: 15840 },
margin: { top: 1224, right: 1224, bottom: 1224, left: 1224 },
}
},
children,
}]
});
}

// ── ROUTES ────────────────────────────────────────────────────────────────────
app.get(’/’, (req, res) => {
res.json({ status: ‘Alliant Agreement Server running’, version: ‘1.0.0’ });
});

app.post(’/generate’, async (req, res) => {
try {
const values = req.body;
if (!values || typeof values !== ‘object’) {
return res.status(400).json({ error: ‘Invalid request body’ });
}
const doc = buildDocument(values);
const buffer = await Packer.toBuffer(doc);
const sellerName = (values.sellerName || ‘Agreement’).replace(/[^a-zA-Z0-9_-]/g, ‘_’);
res.setHeader(‘Content-Type’, ‘application/vnd.openxmlformats-officedocument.wordprocessingml.document’);
res.setHeader(‘Content-Disposition’, `attachment; filename="Alliant_Agreement_${sellerName}.docx"`);
res.setHeader(‘Access-Control-Expose-Headers’, ‘Content-Disposition’);
res.send(buffer);
} catch (err) {
console.error(‘Generation error:’, err);
res.status(500).json({ error: ‘Failed to generate document’, details: err.message });
}
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Alliant Agreement Server listening on port ${PORT}`));
