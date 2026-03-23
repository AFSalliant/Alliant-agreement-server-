var express = require('express');
var cors = require('cors');
var docx = require('docx');

var Document = docx.Document;
var Packer = docx.Packer;
var Paragraph = docx.Paragraph;
var TextRun = docx.TextRun;
var Table = docx.Table;
var TableRow = docx.TableRow;
var TableCell = docx.TableCell;
var AlignmentType = docx.AlignmentType;
var BorderStyle = docx.BorderStyle;
var WidthType = docx.WidthType;
var ShadingType = docx.ShadingType;
var VerticalAlign = docx.VerticalAlign;
var PageBreak = docx.PageBreak;
var UnderlineType = docx.UnderlineType;

var app = express();
app.use(cors());
app.use(express.json({ limit: '2mb' }));

var NAVY = "1E4D78";
var NAVY2 = "16355C";
var NAVY3 = "1D3659";
var BLUE = "4E80BC";
var ROW1 = "E9F0FA";
var ROW2 = "F3F6FA";
var ROW3 = "D8E4F4";
var GRID = "C0D0E8";
var PW = 9360;

function cb(color) {
  return {
    top:    { style: BorderStyle.SINGLE, size: 4, color: color },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: color },
    left:   { style: BorderStyle.SINGLE, size: 4, color: color },
    right:  { style: BorderStyle.SINGLE, size: 4, color: color }
  };
}

var NB = {
  top:    { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right:  { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }
};

var CM = { top: 80, bottom: 80, left: 120, right: 120 };
var LINE = "___________________________";
var LLINE = "____________________________________";

function fm(val) {
  if (!val) return "$" + LINE;
  var n = parseFloat(String(val).replace(/[^0-9.]/g, ""));
  if (isNaN(n)) return "$" + val;
  return "$" + n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function fp(val) { return val ? val + "%" : "_____%" ; }
function dv(val, fb) { return val || (fb !== undefined ? fb : LINE); }

function hRule() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 1 } },
    spacing: { before: 0, after: 100 },
    children: []
  });
}

function sp(pts) {
  return new Paragraph({ spacing: { before: 0, after: pts || 80 }, children: [] });
}

function ctr(text, size, color) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text: text, bold: true, size: size || 22, color: color || NAVY2, font: "Arial" })]
  });
}

function sub(text, size, color) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 30 },
    children: [new TextRun({ text: text, size: size || 19, color: color || BLUE, font: "Arial" })]
  });
}

function sec(num, text) {
  return new Paragraph({
    spacing: { before: 140, after: 50 },
    children: [new TextRun({ text: num + ". " + text, bold: true, size: 19, color: NAVY3, font: "Arial" })]
  });
}

function bp(runs) {
  return new Paragraph({
    alignment: AlignmentType.BOTH,
    spacing: { before: 0, after: 60 },
    children: runs.map(function(r) {
      if (typeof r === "string") return new TextRun({ text: r, size: 19, font: "Arial" });
      return new TextRun(Object.assign({ size: 19, font: "Arial" }, r));
    })
  });
}

function subp(letter, runs) {
  var trs = typeof runs === "string"
    ? [new TextRun({ text: runs, size: 19, font: "Arial" })]
    : runs.map(function(r) { return new TextRun(Object.assign({ size: 19, font: "Arial" }, r)); });
  return new Paragraph({
    alignment: AlignmentType.BOTH,
    spacing: { before: 0, after: 60 },
    indent: { left: 400, hanging: 400 },
    children: [new TextRun({ text: "(" + letter + ")  ", size: 19, font: "Arial" })].concat(trs)
  });
}

function hdrRow(cols, widths) {
  return new TableRow({
    tableHeader: true,
    children: cols.map(function(text, i) {
      return new TableCell({
        borders: cb(NAVY),
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        margins: CM,
        width: { size: widths[i], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ children: [new TextRun({ text: text, bold: true, color: "FFFFFF", size: 18, font: "Arial" })] })]
      });
    })
  });
}

function dRow(cells, widths, shade) {
  return new TableRow({
    children: cells.map(function(cell, i) {
      var runs = Array.isArray(cell)
        ? cell.map(function(r) { return new TextRun(Object.assign({ size: 18, font: "Arial" }, r)); })
        : [new TextRun({ text: String(cell), size: 18, font: "Arial" })];
      return new TableCell({
        borders: cb(GRID),
        shading: { fill: shade, type: ShadingType.CLEAR },
        margins: CM,
        width: { size: widths[i], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ children: runs })]
      });
    })
  });
}

function altRows(rows, widths) {
  return rows.map(function(r, i) { return dRow(r, widths, i % 2 === 0 ? ROW1 : ROW2); });
}

function sigTbl(left, right) {
  var half = Math.floor(PW / 2);
  return new Table({
    width: { size: PW, type: WidthType.DXA },
    columnWidths: [half, PW - half],
    rows: [new TableRow({ children: [
      new TableCell({ borders: NB, width: { size: half, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 0, right: 200 },
        children: [new Paragraph({ children: [
          new TextRun({ text: left, bold: true, size: 19, font: "Arial" }),
          new TextRun({ text: "\n\nSignature:  " + LLINE, size: 19, font: "Arial" }),
          new TextRun({ text: "\n\nName:  " + LLINE, size: 19, font: "Arial" }),
          new TextRun({ text: "\n\nDate:  " + LLINE, size: 19, font: "Arial" })
        ], spacing: { after: 0 } })]
      }),
      new TableCell({ borders: NB, width: { size: PW - half, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 200, right: 0 },
        children: [new Paragraph({ children: [
          new TextRun({ text: right, bold: true, size: 19, font: "Arial" }),
          new TextRun({ text: "\n\nSignature:  " + LLINE, size: 19, font: "Arial" }),
          new TextRun({ text: "\n\nName / Title:  " + LLINE, size: 19, font: "Arial" }),
          new TextRun({ text: "\n\nDate:  " + LLINE, size: 19, font: "Arial" })
        ], spacing: { after: 0 } })]
      })
    ]})]
  });
}

function buildDoc(v) {
  var ch = [];
  var p = function() { for (var i = 0; i < arguments.length; i++) ch.push(arguments[i]); };

  var m = function(id) { return fm(v[id]); };
  var pct = function(id) { return fp(v[id]); };
  var d = function(id, fb) { return dv(v[id], fb); };

  var modeDesc = v.interestMode === "simple"
    ? (v.simpleRate || "___") + "% per six-month Settlement Period applied to the Total Purchase Cost (not compounded)"
    : (v.monthlyRate || "___") + "% per month, compounded monthly on the outstanding balance";

  // HEADER
  p(ctr("ALLIANT LEGAL FUNDING LLC", 26, NAVY2));
  p(sub("Non-Recourse Purchase and Sale Agreement", 19, BLUE));
  p(sub("Consumer Personal Injury Pre-Settlement Funding", 19, BLUE));
  p(sp(60), hRule());

  // Notice box
  p(new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: [PW],
    rows: [new TableRow({ children: [new TableCell({
      borders: cb(NAVY), shading: { fill: ROW2, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 180, right: 180 },
      width: { size: PW, type: WidthType.DXA },
      children: [new Paragraph({ alignment: AlignmentType.BOTH, spacing: { before: 0, after: 0 }, children: [
        new TextRun({ text: "Important Notice: ", bold: true, size: 18, font: "Arial" }),
        new TextRun({ text: "This transaction is a non-recourse purchase and sale of a contingent interest in potential proceeds of a legal claim. It is not a loan. If there is no recovery from the claim, Purchaser receives nothing, except as expressly provided for fraud, material misrepresentation, or other specified default provisions set forth herein.", size: 18, font: "Arial" })
      ]})]
    })]})],
  }));

  p(sp(60));
  p(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: GRID, space: 1 } }, spacing: { before: 0, after: 80 }, children: [] }));
  p(sp(60));

  // TITLE
  p(ctr("NON-RECOURSE PURCHASE AND SALE AGREEMENT", 24, NAVY2));
  p(sp(40));

  p(bp([
    { text: "THIS NON-RECOURSE PURCHASE AND SALE AGREEMENT", bold: true },
    { text: " (the \"Agreement\") is made on " },
    { text: d("agreementDate"), bold: true },
    { text: " (\"Date of Agreement\") by and between " },
    { text: "Alliant Legal Funding LLC", bold: true },
    { text: ", a " + d("purchaserState") + " limited liability company with offices at " + d("purchaserAddress") + " (\"Purchaser\", \"us\") and " },
    { text: d("sellerName"), bold: true },
    { text: ", residing at " + d("sellerStreet") + ", " + d("sellerCityStateZip") + " (\"Seller\", \"you\")." }
  ]));

  p(sp(40));
  p(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: GRID, space: 1 } }, spacing: { before: 0, after: 80 }, children: [] }));
  p(sp(40));

  p(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 80, after: 50 },
    children: [new TextRun({ text: "Purchase and Sale", bold: true, size: 19, color: NAVY3, font: "Arial", underline: { type: UnderlineType.SINGLE } })]
  }));
  p(sp(30));
  p(bp([{ text: "Seller holds the Legal Claim described in Schedule 1. Pursuant to this Agreement, Seller hereby agrees to sell and assign to Purchaser, without recourse, Seller's entire right, title, and interest in the Proceeds of the Legal Claim. As payment, Purchaser agrees to pay Seller the Funded Amount of " }, { text: m("fundedAmount"), bold: true }, { text: " as further described in this Agreement." }]));

  p(sp(40));
  p(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: GRID, space: 1 } }, spacing: { before: 0, after: 60 }, children: [] }));
  p(sp(60));

  // DISCLOSURE TABLE
  p(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "CONSUMER DISCLOSURE AND TRANSACTION SUMMARY", bold: true, size: 21, color: NAVY2, font: "Arial" })] }));

  var dW = [3960, 1800, 3600];
  p(new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: dW,
    rows: [
      hdrRow(["Transaction Item", "Amount", "Notes"], dW)
    ].concat(altRows([
      ["Funded Amount", m("fundedAmount"), "Gross amount advanced by Purchaser"],
      ["Amount Paid Directly to Seller", m("disbursalToSeller"), "Net cash to consumer"],
      ["Amount Paid to Prior Funder(s), if any", m("payoffPriorFunder"), "Leave blank if none"],
      ["Origination / Application Fee", m("originationFee"), "One-time fee, if applicable"],
      ["Administrative / Underwriting Fee", m("adminFee"), "One-time fee, if applicable"],
      ["Annual Maintenance Fee", m("annualMaintFee"), "Charged each year from Date of Agreement"],
      ["Payment Delivery Fee", m("deliveryFee"), "Wire / ACH / check fee, if applicable"],
      ["Total Charges Disclosed at Funding", m("totalCharges"), "Includes all charges known at funding"],
      ["Total Purchase Cost / Base Contracted Amount", m("totalPurchaseCost"), "Funded amount plus disclosed charges"],
      ["Maximum Contracted Amount (if capped)", "Not to exceed " + m("maxContractedAmount"), "If agreement includes a cap"],
      ["Interest Calculation Method", modeDesc, "Applied per Schedule 2"],
      ["APR (disclosure only)", pct("apr"), "For disclosure purposes"]
    ], dW))
  }));

  p(sp(80));

  // Acknowledgment bar
  var aW = [2400, PW - 2400];
  p(new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: aW,
    rows: [new TableRow({ children: [
      new TableCell({ borders: cb(GRID), shading: { fill: ROW3, type: ShadingType.CLEAR }, margins: CM, width: { size: aW[0], type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: "Seller Acknowledgment", bold: true, size: 18, font: "Arial" })] })] }),
      new TableCell({ borders: cb(GRID), shading: { fill: ROW3, type: ShadingType.CLEAR }, margins: CM, width: { size: aW[1], type: WidthType.DXA },
        children: [new Paragraph({ children: [
          new TextRun({ text: "Seller Signature:  ", size: 18, font: "Arial" }),
          new TextRun({ text: LLINE, size: 18, font: "Arial" }),
          new TextRun({ text: "     Date:  ", size: 18, font: "Arial" }),
          new TextRun({ text: LINE, size: 18, font: "Arial" })
        ]})] })
    ]})]
  }));

  p(sp(100));
  p(new Paragraph({ spacing: { before: 100, after: 40 }, children: [new TextRun({ text: "Payment Schedule - Disclosure Summary", bold: true, size: 19, color: NAVY3, font: "Arial" })] }));

  var psW = [2520, 3240, 3600];
  p(new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: psW,
    rows: [
      hdrRow(["Date Range From Funding Date", "Contracted Amount Due if Resolved in This Interval", "Annualized Rate / Disclosure Note"], psW)
    ].concat(altRows([
      ["0 - 6 months", m("disc_0_6"), "For disclosure only"],
      ["6 - 12 months", m("disc_6_12"), "For disclosure only"],
      ["12 - 18 months", m("disc_12_18"), "For disclosure only"],
      ["18 - 24 months", m("disc_18_24"), "For disclosure only"],
      ["24 - 30 months", m("disc_24_30"), "For disclosure only"],
      ["30 months and thereafter", m("disc_30plus"), "For disclosure only"]
    ], psW))
  }));

  p(sp(80));
  p(bp([{ text: "This transaction is a non-recourse purchase of a contingent interest in proceeds of a legal claim. It is not a loan. If Seller complies with the Agreement and there is no recovery from the Legal Claim, Purchaser receives nothing. If the recovery is less than the Contracted Amount, Purchaser is limited to the available proceeds after payment of attorneys fees, court costs, and any liens or priorities that apply by law or by this Agreement.", size: 17 }]));

  p(new Paragraph({ children: [new PageBreak()] }));

  // SECTIONS 1-5
  p(sec("1", "Recitals"));
  p(subp("a", "Seller has asserted a bona fide personal injury or related civil claim described in Schedule 1 (the \"Legal Claim\"). Seller anticipates that the Legal Claim may result in proceeds by settlement, judgment, award, verdict, arbitration, or other resolution."));
  p(subp("b", "Seller desires immediate funds for personal purposes unrelated to the prosecution of the Legal Claim, and Purchaser is willing to purchase from Seller a contingent right to receive a portion of the proceeds of the Legal Claim on the terms stated herein."));
  p(subp("c", "The parties intend this transaction to be a true purchase and sale of a contingent interest in proceeds only, and not a loan, extension of credit, or assignment of the cause of action itself."));

  p(sec("2", "Definitions"));
  p(subp("a", [{ text: "\"Contracted Amount\"", bold: true }, { text: " means the amount that Purchaser is entitled to receive from the proceeds of the Legal Claim, as determined under Schedule 2 as of the date Purchaser is paid." }]));
  p(subp("b", [{ text: "\"Funded Amount\"", bold: true }, { text: " means " + m("fundedAmount") + " advanced by Purchaser to or on behalf of Seller at closing." }]));
  p(subp("c", [{ text: "\"Proceeds\"", bold: true }, { text: " means all money or things of value paid or payable on account of the Legal Claim, whether by settlement, judgment, verdict, or otherwise." }]));
  p(subp("d", [{ text: "\"Purchased Interest\"", bold: true }, { text: " means Purchaser's contingent right to receive the Contracted Amount from the Proceeds." }]));
  p(subp("e", [{ text: "\"Use Fee\"", bold: true }, { text: " means the periodic charge calculated as follows: " + modeDesc + "." }]));
  p(subp("f", [{ text: "\"Settlement Period\"", bold: true }, { text: " means each six-month interval commencing on the Date of Agreement." }]));
  p(subp("g", [{ text: "\"Annual Maintenance Fee\"", bold: true }, { text: " means the recurring annual fee of " + m("annualMaintFee") + " charged each year from the Date of Agreement." }]));

  p(sec("3", "Nature of Transaction; Non-Recourse; No Loan"));
  p(subp("a", "This Agreement is a non-recourse purchase and sale. Purchaser is not making a loan to Seller."));
  p(subp("b", "If the Legal Claim results in no Proceeds, Purchaser receives nothing and Seller owes nothing, except for damages resulting from Seller's fraud, intentional misrepresentation, or conversion of Proceeds."));
  p(subp("c", "If the Legal Claim results in insufficient Proceeds, Seller shall have no further personal liability for any deficiency."));
  p(subp("d", "Nothing in this Agreement transfers the Legal Claim itself, and Purchaser shall have no right to direct litigation or settlement decisions."));
  p(subp("e", "Seller intends this transaction to be a purchase and sale and not a loan."));

  p(sec("4", "Purchase and Sale; Consideration; Purchaser's Acceptance"));
  p(subp("a", "Seller hereby sells, assigns, and transfers to Purchaser the Purchased Interest in exchange for the Funded Amount of " + m("fundedAmount") + "."));
  p(subp("b", "Purchaser's obligations are not effective unless and until Purchaser delivers the Funded Amount. Purchaser retains sole discretion over whether to fund prior to acceptance."));
  p(subp("c", "This Agreement becomes effective only when Purchaser delivers the Funded Amount pursuant to Schedule 5."));
  p(subp("d", "Seller shall not obtain additional funding secured by the same Proceeds without Purchaser's prior written consent."));

  p(sec("5", "Use Fee and Annual Maintenance Fee"));
  p(subp("a", "The Use Fee shall be calculated as follows: " + modeDesc + "."));
  p(subp("b", "The Annual Maintenance Fee of " + m("annualMaintFee") + " shall be added on each anniversary of the Date of Agreement."));
  p(subp("c", "The Contracted Amount equals the Total Purchase Cost plus all accrued Use Fees plus all accrued Annual Maintenance Fees, subject to the Maximum Contracted Amount cap of " + m("maxContractedAmount") + "."));

  p(sec("6", "Payment; Priority; Mechanics of Remittance"));
  p(subp("a", "Upon receipt of any Proceeds, Seller shall pay Purchaser the Contracted Amount from the first Proceeds available after attorneys fees, court costs, and mandatory priority amounts."));
  p(subp("b", "Seller shall not receive any Proceeds until Purchaser has been paid the amount due under Schedule 2."));
  p(subp("c", "If any Proceeds are delivered directly to Seller, Seller shall hold them in trust for Purchaser and deliver same within five (5) business days."));
  p(subp("d", "Seller retains the option to pay the then-current Contracted Amount at any time before resolution of the Legal Claim without any prepayment penalty."));
  p(subp("f", "If Seller proceeds without an attorney, Seller authorizes Purchaser to notify third parties of Purchaser's interest and instructs them to remit Proceeds directly to Purchaser."));

  p(new Paragraph({ children: [new PageBreak()] }));

  p(sec("7", "Seller Representations and Warranties"));
  p(subp("a", "Seller is fully conversant with the English language or requested translated materials, and discussions were conducted accordingly."));
  p(subp("b", "Seller is the lawful owner of the Legal Claim with full authority to enter into this Agreement."));
  p(subp("c", "The Legal Claim was asserted in good faith, and no portion of the Funded Amount will be used to pay litigation expenses."));
  p(subp("d", "All information provided by Seller is true, complete, and not materially misleading."));
  p(subp("e", "Except as disclosed in writing, Seller has not sold, assigned, or encumbered the Legal Claim or any Proceeds."));
  p(subp("f", "There are no outstanding tax liens against Seller."));
  p(subp("g", "Seller has no support obligations constituting a lien against the Proceeds."));
  p(subp("h", "Seller has been advised to review this Agreement with counsel and has done so or expressly waived it."));

  p(sec("8", "Seller Covenants"));
  p(subp("a", "Seller shall cooperate in good faith with the prosecution and resolution of the Legal Claim."));
  p(subp("b", "Seller shall promptly notify Purchaser of any material development affecting the Legal Claim."));
  p(subp("c", "Seller shall not obtain additional funding without Purchaser's prior written consent."));
  p(subp("d", "Seller will provide any new attorney with a signed Irrevocable Letter of Direction."));

  p(sec("9", "No Control Over Claim; No Legal Services"));
  p(subp("a", "Purchaser is not a law firm, is not giving legal advice, and is not assuming any duty to prosecute or manage the Legal Claim."));
  p(subp("b", "All litigation and settlement decisions remain solely with Seller and Seller's Attorney."));

  p(sec("10", "Legal Claim Information Authorization"));
  p(subp("a", "Seller authorizes Purchaser to obtain non-privileged information concerning the Legal Claim, including pleadings, motions, rulings, and periodic status reports, to the extent permitted by law."));

  p(sec("11", "Security Interest; UCC; Protective Measures"));
  p(subp("a", "Seller grants Purchaser a security interest in the Proceeds and authorizes Purchaser to file UCC financing statements describing Purchaser's interest."));

  p(sec("12", "Bankruptcy; Death; Successors"));
  p(subp("a", "In any bankruptcy proceeding, Seller shall describe Purchaser's interest as an asset of Purchaser, not a debt."));
  p(subp("b", "If Seller dies, this Agreement binds Seller's estate solely with respect to the Proceeds."));

  p(sec("13", "Binding Effect; Assignment; Successor Purchaser"));
  p(subp("a", "Purchaser may assign this Agreement without Seller's consent. Seller may not assign this Agreement."));
  p(subp("b", "Upon assignment to a Successor Purchaser, Seller and Seller's Attorney shall direct all payments to Successor Purchaser."));

  p(sec("14", "Events of Default; Remedies; Breach"));
  p(subp("a", "An Event of Default occurs if Seller commits fraud, misrepresentation, diverts Proceeds, or materially breaches this Agreement."));
  p(subp("b", "Seller may be held personally liable for the Contracted Amount in the event of a default."));
  p(subp("c", "Seller shall indemnify and hold Purchaser harmless from losses arising from Seller's default."));

  p(sec("15", "Right of Cancellation"));
  p(subp("a", "Seller may cancel this Agreement without penalty by returning the full Funded Amount of " + m("fundedAmount") + " to Purchaser within five (5) business days after receiving funds."));

  p(sec("16", "Attorney's Fees"));
  p(subp("a", "All of Purchaser's legal fees are Purchaser's sole responsibility. Seller is solely responsible for Seller's own legal fees."));

  p(sec("17", "Governing Law"));
  p(subp("a", "This Agreement shall be governed by the law of " + d("governingLawState") + ". The Arbitration Clause is governed by the Federal Arbitration Act."));
  p(subp("b", "To the extent any claim is not subject to arbitration, it shall be heard in " + d("venueDisputeRes") + "."));

  p(sec("18", "Arbitration Clause"));
  p(sp(40));

  var arbRows2 = [new TableRow({ children: [new TableCell({
    borders: { top: { style: BorderStyle.SINGLE, size: 8, color: NAVY }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.SINGLE, size: 8, color: NAVY }, right: { style: BorderStyle.SINGLE, size: 8, color: NAVY } },
    shading: { fill: ROW3, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 180, right: 180 }, width: { size: PW, type: WidthType.DXA },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ARBITRATION CLAUSE", bold: true, size: 20, font: "Arial" })] })]
  })]})];

  var arbItems = [
    [{ text: "Read this Arbitration Clause carefully. It significantly affects your rights in any dispute with us.", bold: true }],
    [{ text: "* EITHER YOU OR WE MAY CHOOSE TO HAVE ANY DISPUTE BETWEEN US DECIDED BY ARBITRATION AND NOT IN COURT.", bold: true }],
    [{ text: "* IF A DISPUTE IS ARBITRATED, YOU AND WE WILL EACH GIVE UP OUR RIGHT TO A TRIAL BY THE COURT OR A JURY TRIAL.", bold: true }],
    [{ text: "* YOU WILL GIVE UP YOUR RIGHT TO PARTICIPATE AS A CLASS REPRESENTATIVE OR CLASS MEMBER ON ANY CLASS CLAIM YOU MAY HAVE AGAINST US.", bold: true }],
    [{ text: "Governing Law. ", bold: true }, { text: "This Arbitration Clause is governed by the Federal Arbitration Act." }],
    [{ text: "Covered Disputes. ", bold: true }, { text: "The term \"dispute\" includes any claim or dispute between Seller and Purchaser arising from or relating to this Agreement." }],
    [{ text: "Arbitration Procedures. ", bold: true }, { text: "Any dispute will, at either party's election, be resolved by neutral, binding arbitration. Seller expressly waives any right to arbitrate a class action. Seller may choose AAA (www.adr.org) or JAMS (www.jamsadr.com) rules." }],
    [{ text: "Fees. ", bold: true }, { text: "If Seller demands arbitration first, Seller will pay the filing fee if required. The arbitrator's award is final and binding." }],
    [{ text: "Right to Opt Out. ", bold: true }, { text: "Seller may opt out of this Arbitration Clause by sending written notice to Purchaser no later than 45 days after the Date of Agreement." }]
  ];

  arbItems.forEach(function(runs, i) {
    var isLast = i === arbItems.length - 1;
    arbRows2.push(new TableRow({ children: [new TableCell({
      borders: {
        top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        bottom: isLast ? { style: BorderStyle.SINGLE, size: 8, color: NAVY } : { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        left: { style: BorderStyle.SINGLE, size: 8, color: NAVY },
        right: { style: BorderStyle.SINGLE, size: 8, color: NAVY }
      },
      shading: { fill: i % 2 === 0 ? ROW2 : ROW1, type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 180, right: 180 }, width: { size: PW, type: WidthType.DXA },
      children: [new Paragraph({ alignment: AlignmentType.BOTH, spacing: { before: 0, after: 0 },
        children: runs.map(function(r) { return new TextRun(Object.assign({ size: 18, font: "Arial" }, r)); })
      })]
    })]}));
  });

  p(new Table({ width: { size: PW, type: WidthType.DXA }, columnWidths: [PW], rows: arbRows2 }));
  p(sp(80));

  p(sec("19", "Miscellaneous"));
  p(subp("a", "This Agreement constitutes the entire agreement between the parties and supersedes all prior discussions."));
  p(subp("b", "No waiver is effective unless in writing."));
  p(subp("c", "If any provision is held invalid, the remaining provisions remain in effect."));
  p(subp("d", "This Agreement may be executed in counterparts and by electronic signature."));
  p(subp("e", "This Agreement may not be modified orally."));

  p(sp(80), hRule(), sp(60));
  p(bp([
    { text: "DO NOT SIGN THIS AGREEMENT IF IT CONTAINS BLANK SPACES. ", bold: true },
    { text: "IN WITNESS WHEREOF", bold: true },
    { text: ", Seller hereby executes this Agreement." }
  ]));
  p(sp(120));
  p(sigTbl("SELLER / CLAIMANT", "PURCHASER - ALLIANT LEGAL FUNDING LLC"));

  p(new Paragraph({ children: [new PageBreak()] }));

  // SCHEDULE 1
  p(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "SCHEDULE 1", bold: true, size: 21, color: NAVY2, font: "Arial" })] }));
  p(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "Claim Information", bold: true, size: 21, color: NAVY2, font: "Arial" })] }));
  p(hRule(), sp(60));

  var s1W = [3240, PW - 3240];
  p(new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: s1W,
    rows: altRows([
      [[{ text: "Seller Name", bold: true }], d("sellerName")],
      [[{ text: "Date of Loss / Accident", bold: true }], d("dateOfLoss")],
      [[{ text: "Court / Venue", bold: true }], d("courtVenue")],
      [[{ text: "Index / Case Number", bold: true }], d("indexCaseNumber")],
      [[{ text: "Caption", bold: true }], d("claimCaption")],
      [[{ text: "Defendant(s) / Carrier(s)", bold: true }], d("defendants")],
      [[{ text: "Adjuster or Claim Number", bold: true }], d("adjusterClaimNum")],
      [[{ text: "Seller's Attorney / Firm", bold: true }], d("attorneyName") + ", " + d("attorneyFirm")],
      [[{ text: "Known Liens / Prior Funding / Notes", bold: true }], d("knownLiens", "None")]
    ], s1W)
  }));

  p(new Paragraph({ children: [new PageBreak()] }));

  // SCHEDULE 2
  p(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "SCHEDULE 2", bold: true, size: 21, color: NAVY2, font: "Arial" })] }));
  p(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "Payment Schedule and Pricing Table", bold: true, size: 21, color: NAVY2, font: "Arial" })] }));
  p(hRule(), sp(60));
  p(bp([{ text: "Interest Calculation Method: " + modeDesc }]));
  p(sp(40));

  var s2W = [2520, 2880, 3960];
  p(new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: s2W,
    rows: [hdrRow(["Interval", "Contracted Amount Due", "Method / Note"], s2W)].concat(altRows([
      ["0 - 6 months", m("s2_0_6"), "Per interest method above"],
      ["6 - 12 months", m("s2_6_12"), "Per interest method above"],
      ["12 - 18 months", m("s2_12_18"), "Per interest method above"],
      ["18 - 24 months", m("s2_18_24"), "Per interest method above"],
      ["24 - 30 months", m("s2_24_30"), "Per interest method above"],
      ["30 - 36 months", m("s2_30_36"), "Per interest method above"],
      ["Thereafter", m("s2_thereafter"), "Not to exceed cap"]
    ], s2W))
  }));
  p(sp(80));

  var s2eW = [3960, PW - 3960];
  p(new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: s2eW,
    rows: altRows([
      [[{ text: "APR (disclosure only)", bold: true }], pct("apr")],
      [[{ text: "Annual Maintenance Fee", bold: true }], m("annualMaintFee") + " per year"],
      [[{ text: "Maximum Contracted Amount Cap", bold: true }], m("maxContractedAmount")],
      [[{ text: "Governing Law", bold: true }], "State of " + d("governingLawState")],
      [[{ text: "Venue / Dispute Resolution", bold: true }], d("venueDisputeRes")]
    ], s2eW)
  }));

  p(new Paragraph({ children: [new PageBreak()] }));

  // SCHEDULE 3
  p(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "SCHEDULE 3", bold: true, size: 21, color: NAVY2, font: "Arial" })] }));
  p(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "Irrevocable Letter of Direction to Counsel", bold: true, size: 21, color: NAVY2, font: "Arial" })] }));
  p(hRule(), sp(60));
  p(bp([{ text: "To: " + d("attorneyName") + ", Esq. / Firm: " + d("attorneyFirm") }]));
  p(bp([{ text: d("attorneyAddress") }]));
  p(sp(40));
  p(bp([{ text: "I, " }, { text: d("sellerName"), bold: true }, { text: " (\"Seller\"), irrevocably direct you, and any substitute, successor, or additional counsel representing me in the Legal Claim described in Schedule 1, to protect and satisfy the Purchased Interest of Alliant Legal Funding LLC under the Non-Recourse Purchase and Sale Agreement dated " + d("agreementDate") + " (\"Agreement\")." }]));
  p(sp(30));
  p(bp([{ text: "You are directed to: (1) cooperate with Purchaser on non-privileged case information; (2) notify Purchaser upon resolution; (3) remit the Contracted Amount to Purchaser before distributing remaining Proceeds to me; (4) hold Proceeds in trust if disputed; (5) not permit additional funding without consent; (6) notify Purchaser within 48 hours of withdrawal from representation. This direction binds any subsequent attorney until Purchaser confirms payment in full." }]));
  p(sp(80));
  p(sigTbl("SELLER", "ACKNOWLEDGED BY SELLER'S ATTORNEY"));

  p(new Paragraph({ children: [new PageBreak()] }));

  // SCHEDULE 4
  p(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "SCHEDULE 4", bold: true, size: 21, color: NAVY2, font: "Arial" })] }));
  p(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "Attorney Acknowledgment and Undertaking", bold: true, size: 21, color: NAVY2, font: "Arial" })] }));
  p(hRule(), sp(60));
  p(bp([{ text: "I, " }, { text: d("attorneyName"), bold: true }, { text: ", of " }, { text: d("attorneyFirm"), bold: true }, { text: ", acknowledge receipt of the Agreement and agree as follows:" }]));
  p(sp(40));

  var attyList = [
    "I represent Seller in the Legal Claim identified in Schedule 1.",
    "My fee arrangement is contingent and I will disburse Proceeds through my attorney trust account.",
    "I shall not be paid commissions or referral fees related to this Agreement and I have no financial interest in Purchaser.",
    "I am following the instructions in Seller's Irrevocable Letter of Direction.",
    "I will not disburse Proceeds until Purchaser's Contracted Amount is satisfied.",
    "I will contact Purchaser within 10 days of resolution to request a payoff amount.",
    "I will not accept direction from Purchaser on litigation strategy or settlement decisions.",
    "Marking a check \"in full satisfaction\" will not have legal effect without Purchaser's written confirmation.",
    "I will not participate in any future funding without prior written approval of Purchaser and Seller.",
    "If I no longer represent Seller, I will promptly notify Purchaser.",
    "I acknowledge that Purchaser has relied upon this Acknowledgment in providing funding."
  ];

  attyList.forEach(function(item, i) { p(subp(i + 1, item)); });

  p(sp(80));
  p(bp([{ text: "Known liens / prior funding: " + d("knownLiens", "None known") }]));
  p(sp(40));
  p(bp([{ text: "Email: " + d("attorneyEmail") + "     Tel: " + d("attorneyTel") }]));
  p(sp(120));
  p(sigTbl("SELLER'S ATTORNEY / LAW FIRM", "PURCHASER ACKNOWLEDGMENT"));

  p(new Paragraph({ children: [new PageBreak()] }));

  // SCHEDULE 5
  p(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "SCHEDULE 5", bold: true, size: 21, color: NAVY2, font: "Arial" })] }));
  p(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "Funding Disbursement - Fee Schedule and Payment Instructions", bold: true, size: 21, color: NAVY2, font: "Arial" })] }));
  p(hRule(), sp(60));

  p(new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: "Disbursement Fee Schedule", bold: true, size: 19, color: NAVY3, font: "Arial" })] }));

  var fW = [2160, 900, 1800, 4500];
  p(new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: fW,
    rows: [hdrRow(["Payment Method", "Fee", "Estimated Receipt", "Notes"], fW)].concat(altRows([
      ["ACH / Wire Transfer", "$30.00", "Next business day", "Not responsible for holds at receiving bank."],
      ["Check - Express", "$15.00", "2 business days", "No FedEx to P.O. Boxes."],
      ["Check - Overnight", "$30.00", "Next business day", "No FedEx to P.O. Boxes."],
      ["Check - Saturday", "$45.00", "Saturday", "No FedEx to P.O. Boxes."],
      ["Check - USPS", "Free", "5-7 business days", "No guarantee of postal timing."],
      ["Zelle (if offered)", "None", "Same day or next business day", "Transfers are final once sent."]
    ], fW))
  }));

  p(sp(100));
  p(new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: "Payment Method Selection", bold: true, size: 19, color: NAVY3, font: "Arial" })] }));
  p(bp([{ text: "Complete only one section below. Client completes this section at the time of signing." }]));
  p(sp(60));
  p(bp([{ text: "[ ] ACH / Wire Transfer", bold: true }]));

  var achW = [3600, PW - 3600];
  p(new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: achW,
    rows: altRows([
      [[{ text: "Name of Account Holder", bold: true }], v.achAccountHolder || LINE],
      [[{ text: "Bank Name", bold: true }], v.achBankName || LINE],
      [[{ text: "Bank ABA / Routing Number (9 digits only)", bold: true }], v.achRoutingNumber || LINE],
      [[{ text: "Bank Account Number", bold: true }], v.achAccountNumber || LINE],
      [[{ text: "Account Type", bold: true }], v.achAccountType ? v.achAccountType : "[ ] Checking     [ ] Savings"]
    ], achW)
  }));

  p(sp(80));
  p(bp([{ text: "[ ] Paper Check", bold: true }]));

  var chkW = [3600, PW - 3600];
  p(new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: chkW,
    rows: altRows([
      [[{ text: "Issue Check To", bold: true }], v.checkPayee || LINE],
      [[{ text: "Street Address", bold: true }], v.checkStreet || LINE],
      [[{ text: "City, State, Zip Code", bold: true }], v.checkCityStateZip || LINE],
      [[{ text: "Delivery Method", bold: true }], v.checkDelivery || "[ ] Express   [ ] Overnight   [ ] Saturday   [ ] USPS"]
    ], chkW)
  }));

  p(sp(100));
  p(new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: "Seller Personal Information", bold: true, size: 19, color: NAVY3, font: "Arial" })] }));

  var piW = [3240, PW - 3240];
  p(new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: piW,
    rows: altRows([
      [[{ text: "Social Security Number", bold: true }], v.sellerSSN || LINE],
      [[{ text: "Mailing Address", bold: true }], v.sellerMailingAddr || LINE],
      [[{ text: "City, State, Zip Code", bold: true }], v.sellerMailCSZ || LINE],
      [[{ text: "Telephone Number", bold: true }], v.sellerPhone || LINE],
      [[{ text: "Email Address", bold: true }], v.sellerEmail || LINE]
    ], piW)
  }));

  p(sp(60));
  p(bp([{ text: "Please attach a copy of your driver's license or government-issued state ID to this Agreement.", bold: true }]));
  p(sp(60));
  p(bp([{ text: "By signing below, Seller authorizes Alliant Legal Funding LLC and its successors and assigns to disburse the Funded Amount of " + m("fundedAmount") + " through the selected method above.", size: 17 }]));
  p(sp(100));

  var finW = Math.floor(PW / 2);
  p(new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: [finW, PW - finW],
    rows: [new TableRow({ children: [
      new TableCell({ borders: NB, width: { size: finW, type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 0, right: 200 },
        children: [new Paragraph({ children: [
          new TextRun({ text: "Seller Signature:", bold: true, size: 19, font: "Arial" }),
          new TextRun({ text: "\n\n" + LLINE + "\n\nDate:  " + LLINE, size: 19, font: "Arial" })
        ], spacing: { after: 0 } })]
      }),
      new TableCell({ borders: NB, width: { size: PW - finW, type: WidthType.DXA }, margins: { top: 0, bottom: 0, left: 200, right: 0 },
        children: [new Paragraph({ children: [
          new TextRun({ text: "Purchaser Approval:", bold: true, size: 19, font: "Arial" }),
          new TextRun({ text: "\n\n" + LLINE + "\n\nFunding Date:  " + LLINE, size: 19, font: "Arial" })
        ], spacing: { after: 0 } })]
      })
    ]})]
  }));

  return new Document({
    styles: { default: { document: { run: { font: "Arial", size: 19 } }, paragraph: { spacing: { line: 276, lineRule: "auto" } } } },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1224, right: 1224, bottom: 1224, left: 1224 } } },
      children: ch
    }]
  });
}

// ROUTES
app.get("/", function(req, res) {
  res.json({ status: "Alliant Agreement Server running", version: "2.0.0" });
});

app.post("/generate", function(req, res) {
  var values = req.body;
  if (!values || typeof values !== "object") {
    return res.status(400).json({ error: "Invalid request body" });
  }
  buildDoc(values).then ? buildDoc(values) : null;
  var doc = buildDoc(values);
  Packer.toBuffer(doc).then(function(buffer) {
    var sellerName = (values.sellerName || "Agreement").replace(/[^a-zA-Z0-9_-]/g, "_");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", "attachment; filename=\"Alliant_Agreement_" + sellerName + ".docx\"");
    res.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
    res.send(buffer);
  }).catch(function(err) {
    console.error("Generation error:", err);
    res.status(500).json({ error: "Failed to generate document", details: err.message });
  });
});

var PORT = process.env.PORT || 3000;
app.listen(PORT, function() {
  console.log("Alliant Agreement Server listening on port " + PORT);
});
