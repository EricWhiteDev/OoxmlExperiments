// Generate a ~100 page legal document in Word XML Document (flat OPC) format

const fs = require("fs");
const path = require("path");

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

function escapeXml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function wParagraph(text, opts = {}) {
  const { bold, size, heading, spacing, keepNext, indent, numbering } = opts;
  let pPr = "";
  let rPr = "";

  const pPrParts = [];
  const rPrParts = [];

  if (heading) {
    pPrParts.push(`<w:pStyle w:val="Heading${heading}"/>`);
  }
  if (keepNext) {
    pPrParts.push(`<w:keepNext/>`);
  }
  if (spacing) {
    pPrParts.push(`<w:spacing w:before="${spacing.before || 0}" w:after="${spacing.after || 200}" w:line="${spacing.line || 276}" w:lineRule="auto"/>`);
  }
  if (indent) {
    pPrParts.push(`<w:ind w:left="${indent}"/>`);
  }
  if (numbering) {
    pPrParts.push(`<w:numPr><w:ilvl w:val="${numbering.level || 0}"/><w:numId w:val="${numbering.numId || 1}"/></w:numPr>`);
  }
  if (bold) rPrParts.push(`<w:b/>`);
  if (size) rPrParts.push(`<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`);

  if (rPrParts.length > 0) {
    rPr = `<w:rPr>${rPrParts.join("")}</w:rPr>`;
    pPrParts.push(rPr);
  }

  if (pPrParts.length > 0) {
    pPr = `<w:pPr>${pPrParts.join("")}</w:pPr>`;
  }

  return `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function pageBreak() {
  return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
}

// Legal content generators
const legalPhrases = [
  "Subject to the terms and conditions set forth herein, the parties agree to be bound by the provisions of this Agreement.",
  "Notwithstanding anything to the contrary contained in this Agreement, neither party shall be liable for any indirect, incidental, special, consequential, or punitive damages.",
  "Each party represents and warrants that it has full power and authority to enter into this Agreement and to perform its obligations hereunder.",
  "The obligations of confidentiality set forth in this Section shall survive the termination or expiration of this Agreement for a period of five (5) years.",
  "This Agreement shall be governed by and construed in accordance with the laws of the State of New York, without regard to its conflict of laws principles.",
  "Any dispute arising out of or relating to this Agreement shall be resolved through binding arbitration in accordance with the rules of the American Arbitration Association.",
  "Neither party may assign or transfer this Agreement, in whole or in part, without the prior written consent of the other party, except in connection with a merger, acquisition, or sale of all or substantially all of its assets.",
  "This Agreement, together with all exhibits and schedules attached hereto, constitutes the entire agreement between the parties with respect to the subject matter hereof and supersedes all prior and contemporaneous agreements, understandings, negotiations, and discussions.",
  "No amendment, modification, or waiver of any provision of this Agreement shall be effective unless in writing and signed by both parties.",
  "If any provision of this Agreement is held to be invalid, illegal, or unenforceable, the remaining provisions shall continue in full force and effect.",
  "The failure of either party to enforce any provision of this Agreement shall not constitute a waiver of such party's right to enforce such provision or any other provision in the future.",
  "All notices required or permitted under this Agreement shall be in writing and shall be deemed given when delivered personally, sent by confirmed electronic transmission, or sent by certified or registered mail, return receipt requested.",
  "Each party shall indemnify, defend, and hold harmless the other party and its officers, directors, employees, agents, and affiliates from and against any and all claims, damages, losses, costs, and expenses arising out of or relating to any breach of this Agreement.",
  "The parties acknowledge that the remedies at law for any breach of the obligations under this Agreement would be inadequate and that the non-breaching party shall be entitled to seek equitable relief, including injunction and specific performance.",
  "This Agreement may be executed in counterparts, each of which shall be deemed an original, but all of which together shall constitute one and the same instrument.",
  "The headings and captions in this Agreement are for convenience of reference only and shall not affect the interpretation or construction of this Agreement.",
  "No third party shall be deemed a beneficiary of this Agreement, and nothing contained herein shall be construed to create any rights enforceable by any person or entity not a party to this Agreement.",
  "The waiver by either party of a breach of any provision of this Agreement shall not operate or be construed as a waiver of any subsequent breach of the same or any other provision.",
  "Each party shall bear its own costs and expenses incurred in connection with the negotiation, preparation, execution, and performance of this Agreement.",
  "Force Majeure: Neither party shall be liable for any failure or delay in performing its obligations under this Agreement to the extent that such failure or delay results from circumstances beyond the reasonable control of that party.",
];

const definitions = [
  ['"Affiliate"', 'means any entity that directly or indirectly controls, is controlled by, or is under common control with a party, where "control" means the ownership of more than fifty percent (50%) of the voting securities or equivalent ownership interest.'],
  ['"Business Day"', "means any day other than a Saturday, Sunday, or a day on which banking institutions in New York, New York are authorized or required by law or executive order to remain closed."],
  ['"Change of Control"', "means any merger, consolidation, reorganization, or sale of all or substantially all of the assets of a party, or any transaction or series of related transactions resulting in a change in the beneficial ownership of more than fifty percent (50%) of the outstanding voting securities of a party."],
  ['"Claim"', "means any claim, demand, suit, action, proceeding, investigation, or inquiry, whether civil, criminal, administrative, or otherwise, and whether at law or in equity."],
  ['"Confidential Information"', "means all non-public information, whether written, oral, electronic, or visual, disclosed by one party to the other party in connection with this Agreement, including but not limited to trade secrets, business plans, financial information, customer data, technical specifications, and proprietary methodologies."],
  ['"Damages"', "means any and all losses, damages, liabilities, costs, and expenses, including reasonable attorneys' fees and court costs."],
  ['"Deliverables"', "means all work product, materials, documents, data, and other tangible or intangible items to be provided by a party under this Agreement."],
  ['"Effective Date"', "means the date on which this Agreement is executed by the last party to sign, as indicated on the signature page hereof."],
  ['"Governmental Authority"', "means any federal, state, local, or foreign government, or any court, administrative agency, regulatory body, commission, or other governmental authority or instrumentality."],
  ['"Indemnified Party"', "means the party seeking indemnification under Article VIII of this Agreement."],
  ['"Indemnifying Party"', "means the party from whom indemnification is sought under Article VIII of this Agreement."],
  ['"Intellectual Property"', "means all patents, patent applications, trademarks, trademark applications, service marks, trade names, copyrights, trade secrets, know-how, inventions, processes, formulae, algorithms, software, data, and all other intellectual property rights and proprietary information."],
  ['"Law"', "means any statute, law, ordinance, regulation, rule, code, order, constitution, treaty, common law, judgment, decree, or other requirement of any Governmental Authority."],
  ['"Material Adverse Effect"', "means any event, change, circumstance, occurrence, or development that, individually or in the aggregate, has had or would reasonably be expected to have a material adverse effect on the business, assets, liabilities, financial condition, or results of operations of a party."],
  ['"Permitted Use"', "means the use of the Deliverables solely for the purposes contemplated by this Agreement and in accordance with the terms and conditions set forth herein."],
  ['"Person"', "means any individual, corporation, partnership, limited liability company, joint venture, association, trust, unincorporated organization, governmental authority, or other entity."],
  ['"Representatives"', "means, with respect to a party, such party's officers, directors, employees, agents, advisors, consultants, and other representatives."],
  ['"Term"', "means the period commencing on the Effective Date and continuing until the earlier of (a) the date of termination of this Agreement in accordance with Article X, or (b) the fifth (5th) anniversary of the Effective Date."],
  ['"Territory"', "means the United States of America and its territories and possessions, the European Union member states, the United Kingdom, Canada, Australia, and Japan."],
  ['"Transaction"', "means the transactions contemplated by this Agreement, including the delivery of Deliverables and the performance of services as described in the Exhibits attached hereto."],
];

const sections = [
  {
    title: "RECITALS",
    subsections: [
      { title: null, paragraphs: [
        "WHEREAS, Party A is a corporation duly organized and existing under the laws of the State of Delaware, with its principal place of business located at 123 Corporate Boulevard, Suite 500, New York, New York 10001, and is engaged in the business of providing professional consulting services, technology solutions, and related advisory services to corporate clients throughout the Territory;",
        "WHEREAS, Party B is a limited liability company duly organized and existing under the laws of the State of California, with its principal place of business located at 456 Innovation Drive, San Francisco, California 94105, and is engaged in the business of developing, licensing, and maintaining proprietary software applications and technology platforms;",
        "WHEREAS, Party A desires to retain Party B to provide certain services and deliverables as more particularly described in the Exhibits attached hereto, and Party B desires to provide such services and deliverables, subject to the terms and conditions of this Agreement;",
        "WHEREAS, the parties have conducted extensive negotiations and due diligence and have determined that entering into this Agreement is in the mutual best interests of both parties and their respective shareholders, members, and stakeholders;",
        "NOW, THEREFORE, in consideration of the mutual covenants, agreements, representations, and warranties contained herein, and for other good and valuable consideration, the receipt and sufficiency of which are hereby acknowledged, the parties agree as follows:",
      ]},
    ],
  },
  {
    title: "ARTICLE I - DEFINITIONS",
    subsections: [
      { title: "Section 1.1. Defined Terms", paragraphs: definitions.map(([term, def]) => `${term} ${def}`) },
      { title: "Section 1.2. Interpretation", paragraphs: [
        "Unless the context otherwise requires, (a) words in the singular include the plural and vice versa, (b) words importing any gender include all genders, (c) references to 'including' or 'includes' shall mean 'including without limitation' or 'includes without limitation,' (d) references to any agreement, document, or instrument mean such agreement, document, or instrument as amended, supplemented, or modified from time to time, and (e) references to any Law mean such Law as amended, supplemented, or modified from time to time, and include any successor legislation thereto and any regulations promulgated thereunder.",
        "The table of contents and headings contained in this Agreement are for reference purposes only and shall not affect in any way the meaning or interpretation of this Agreement. All references to Articles, Sections, Exhibits, and Schedules shall be deemed references to Articles and Sections of, and Exhibits and Schedules to, this Agreement unless the context otherwise requires.",
      ]},
    ],
  },
  {
    title: "ARTICLE II - SCOPE OF SERVICES",
    subsections: [
      { title: "Section 2.1. Engagement", paragraphs: [
        "Party A hereby engages Party B, and Party B hereby accepts such engagement, to provide the services described in Exhibit A attached hereto (the 'Services') during the Term of this Agreement. Party B shall perform the Services in a professional and workmanlike manner, consistent with industry standards and best practices.",
        "Party B shall devote such time, attention, and resources as are reasonably necessary to perform the Services in accordance with the specifications, timelines, and milestones set forth in Exhibit A. Party B shall assign qualified personnel with appropriate skills, experience, and expertise to perform the Services.",
      ]},
      { title: "Section 2.2. Deliverables", paragraphs: [
        "In connection with the performance of the Services, Party B shall provide to Party A the Deliverables described in Exhibit B attached hereto, in accordance with the delivery schedule and acceptance criteria set forth therein.",
        "All Deliverables shall conform to the specifications, requirements, and quality standards set forth in Exhibit B. Party A shall have the right to review and test each Deliverable within thirty (30) Business Days following delivery to determine whether such Deliverable conforms to the applicable specifications and acceptance criteria.",
        "If Party A determines that a Deliverable does not conform to the applicable specifications or acceptance criteria, Party A shall provide written notice to Party B specifying the deficiencies in reasonable detail. Party B shall, at its own cost and expense, correct such deficiencies and redeliver the corrected Deliverable within fifteen (15) Business Days following receipt of such notice.",
      ]},
      { title: "Section 2.3. Change Orders", paragraphs: [
        "Either party may request changes to the scope of the Services or Deliverables by submitting a written change order request to the other party. Each change order request shall describe the proposed changes in reasonable detail, including the anticipated impact on the project timeline, budget, and resources.",
        "No change to the scope of the Services or Deliverables shall be effective unless and until a written change order has been executed by authorized representatives of both parties. The parties shall negotiate in good faith regarding the terms of any proposed change order, including any adjustments to the fees, timeline, or other terms of this Agreement.",
      ]},
      { title: "Section 2.4. Project Governance", paragraphs: [
        "The parties shall establish a joint project steering committee (the 'Steering Committee') consisting of two (2) representatives designated by each party. The Steering Committee shall meet at least monthly, or more frequently as necessary, to review the progress of the Services, discuss any issues or concerns, and make decisions regarding the project.",
        "Each party shall designate a project manager who shall serve as the primary point of contact for all day-to-day communications regarding the Services. The project managers shall be responsible for coordinating the activities of the parties, monitoring progress against the project plan, and escalating issues to the Steering Committee as appropriate.",
        "Party B shall provide Party A with written progress reports on a bi-weekly basis, detailing the status of the Services and Deliverables, any issues or risks identified, and the proposed mitigation strategies. Party B shall also provide such additional reports and information as Party A may reasonably request from time to time.",
      ]},
    ],
  },
  {
    title: "ARTICLE III - COMPENSATION AND PAYMENT",
    subsections: [
      { title: "Section 3.1. Fees", paragraphs: [
        "In consideration of the Services to be performed and the Deliverables to be provided by Party B under this Agreement, Party A shall pay Party B the fees set forth in Exhibit C attached hereto (the 'Fees'). The Fees shall be the sole and exclusive compensation payable to Party B for the performance of the Services and the delivery of the Deliverables, unless otherwise agreed in writing by the parties.",
        "Party B shall invoice Party A on a monthly basis for the Fees accrued during the preceding calendar month, together with reasonable supporting documentation. Each invoice shall itemize the Services performed, the hours expended by each Party B personnel, and any reimbursable expenses incurred during the applicable period.",
      ]},
      { title: "Section 3.2. Payment Terms", paragraphs: [
        "Party A shall pay each undisputed invoice within forty-five (45) days following receipt of such invoice. All payments shall be made in United States dollars by wire transfer of immediately available funds to the bank account designated by Party B in writing.",
        "If Party A disputes any portion of an invoice, Party A shall notify Party B in writing within fifteen (15) days following receipt of such invoice, specifying the disputed amount and the basis for the dispute in reasonable detail. The parties shall work together in good faith to resolve any such dispute within thirty (30) days following the date of Party A's notice. Party A shall pay the undisputed portion of any invoice in accordance with the payment terms set forth above.",
        "Any undisputed amounts not paid when due shall bear interest at the rate of one and one-half percent (1.5%) per month, or the maximum rate permitted by applicable Law, whichever is less, from the due date until the date of payment.",
      ]},
      { title: "Section 3.3. Expenses", paragraphs: [
        "Party A shall reimburse Party B for all reasonable and necessary out-of-pocket expenses incurred by Party B in connection with the performance of the Services, provided that (a) such expenses are pre-approved in writing by Party A, (b) Party B provides receipts or other reasonable documentation supporting such expenses, and (c) such expenses comply with Party A's expense reimbursement policies as communicated to Party B from time to time.",
        "Travel expenses shall be reimbursed in accordance with Party A's travel policy. Air travel shall be at coach class rates unless otherwise approved by Party A. Hotel accommodations shall not exceed the applicable per diem rates established by the United States General Services Administration for the applicable location.",
      ]},
      { title: "Section 3.4. Taxes", paragraphs: [
        "The Fees are exclusive of all sales, use, value-added, goods and services, withholding, and other taxes, levies, duties, and assessments of any kind imposed by any Governmental Authority (collectively, 'Taxes'). Party A shall be responsible for all Taxes arising from or relating to the transactions contemplated by this Agreement, other than taxes based on Party B's net income.",
        "Each party shall provide the other party with such tax forms, certificates, and other documentation as may be reasonably required to establish any available exemption from, or reduction of, any applicable Taxes.",
      ]},
    ],
  },
  {
    title: "ARTICLE IV - INTELLECTUAL PROPERTY",
    subsections: [
      { title: "Section 4.1. Ownership of Deliverables", paragraphs: [
        "Subject to Section 4.2, all Deliverables created by Party B in the performance of the Services shall be considered 'works made for hire' as that term is defined under the United States Copyright Act, and all right, title, and interest in and to such Deliverables, including all Intellectual Property rights therein, shall vest exclusively in Party A upon creation.",
        "To the extent that any Deliverable or any portion thereof does not qualify as a work made for hire, Party B hereby irrevocably assigns, transfers, and conveys to Party A all right, title, and interest in and to such Deliverable, including all Intellectual Property rights therein, free and clear of all liens, encumbrances, and claims of any kind.",
        "Party B shall execute and deliver such documents and instruments, and take such further actions, as Party A may reasonably request to evidence, perfect, or protect Party A's ownership rights in the Deliverables.",
      ]},
      { title: "Section 4.2. Party B Pre-Existing IP", paragraphs: [
        "Notwithstanding Section 4.1, Party B shall retain all right, title, and interest in and to any Intellectual Property that (a) was developed by Party B prior to the Effective Date, (b) is developed by Party B independently of the Services and without reference to Party A's Confidential Information, or (c) constitutes general knowledge, skills, or experience possessed by Party B's personnel (collectively, 'Party B Pre-Existing IP').",
        "To the extent that any Party B Pre-Existing IP is incorporated into or is necessary for the use of any Deliverable, Party B hereby grants to Party A a non-exclusive, perpetual, irrevocable, worldwide, fully paid-up, royalty-free license, with the right to sublicense, to use, reproduce, modify, create derivative works of, distribute, display, and perform such Party B Pre-Existing IP solely as incorporated in or necessary for the use of the Deliverables.",
      ]},
      { title: "Section 4.3. License to Party A Materials", paragraphs: [
        "Party A hereby grants to Party B a limited, non-exclusive, non-transferable license to use Party A's materials, data, and Intellectual Property solely to the extent necessary for Party B to perform the Services during the Term. This license shall terminate automatically upon the expiration or termination of this Agreement.",
        "Party B shall not use Party A's materials, data, or Intellectual Property for any purpose other than the performance of the Services, and shall not disclose, distribute, or make available Party A's materials, data, or Intellectual Property to any third party without Party A's prior written consent.",
      ]},
    ],
  },
  {
    title: "ARTICLE V - CONFIDENTIALITY",
    subsections: [
      { title: "Section 5.1. Obligations of Confidentiality", paragraphs: [
        "Each party (the 'Receiving Party') agrees that during the Term and for a period of five (5) years following the expiration or termination of this Agreement, it shall (a) hold in strict confidence all Confidential Information of the other party (the 'Disclosing Party'), (b) not disclose any Confidential Information to any third party, except as expressly permitted herein, and (c) not use any Confidential Information for any purpose other than the performance of its obligations or the exercise of its rights under this Agreement.",
        "The Receiving Party shall limit access to the Disclosing Party's Confidential Information to those of its Representatives who have a need to know such information in connection with this Agreement and who are bound by confidentiality obligations at least as protective as those set forth in this Article V.",
        "The Receiving Party shall protect the Disclosing Party's Confidential Information using the same degree of care that it uses to protect its own confidential information of like nature, but in no event less than a reasonable degree of care.",
      ]},
      { title: "Section 5.2. Exceptions", paragraphs: [
        "The obligations of confidentiality set forth in Section 5.1 shall not apply to any information that (a) is or becomes generally available to the public other than as a result of a disclosure by the Receiving Party or its Representatives in violation of this Agreement, (b) was available to the Receiving Party on a non-confidential basis prior to its disclosure by the Disclosing Party, (c) becomes available to the Receiving Party on a non-confidential basis from a source other than the Disclosing Party or its Representatives, provided that such source is not known by the Receiving Party to be bound by a confidentiality agreement with or other obligation of secrecy to the Disclosing Party, or (d) is independently developed by the Receiving Party without reference to or use of the Disclosing Party's Confidential Information.",
      ]},
      { title: "Section 5.3. Compelled Disclosure", paragraphs: [
        "If the Receiving Party or any of its Representatives is compelled by applicable Law or legal process to disclose any Confidential Information of the Disclosing Party, the Receiving Party shall (a) provide the Disclosing Party with prompt written notice of such requirement so that the Disclosing Party may seek a protective order or other appropriate remedy, and (b) cooperate with the Disclosing Party, at the Disclosing Party's expense, in seeking such protective order or other remedy.",
        "If such protective order or other remedy is not obtained, the Receiving Party shall disclose only that portion of the Confidential Information that is legally required to be disclosed and shall use commercially reasonable efforts to ensure that confidential treatment is accorded to such Confidential Information.",
      ]},
      { title: "Section 5.4. Return or Destruction", paragraphs: [
        "Upon the expiration or termination of this Agreement, or upon the written request of the Disclosing Party at any time, the Receiving Party shall promptly (a) return to the Disclosing Party all tangible materials containing or embodying the Disclosing Party's Confidential Information, and (b) destroy all electronic copies of the Disclosing Party's Confidential Information in the Receiving Party's possession or control, and certify such destruction in writing to the Disclosing Party.",
        "Notwithstanding the foregoing, the Receiving Party may retain one (1) archival copy of the Disclosing Party's Confidential Information solely for the purpose of determining its obligations under this Article V, provided that such archival copy is maintained in a secure location with restricted access.",
      ]},
    ],
  },
  {
    title: "ARTICLE VI - REPRESENTATIONS AND WARRANTIES",
    subsections: [
      { title: "Section 6.1. Mutual Representations and Warranties", paragraphs: [
        "Each party represents and warrants to the other party that: (a) it is duly organized, validly existing, and in good standing under the laws of its jurisdiction of organization; (b) it has full corporate or organizational power and authority to execute, deliver, and perform this Agreement; (c) the execution, delivery, and performance of this Agreement have been duly authorized by all necessary corporate or organizational action; (d) this Agreement constitutes a legal, valid, and binding obligation of such party, enforceable against it in accordance with its terms, subject to applicable bankruptcy, insolvency, reorganization, moratorium, and other laws affecting creditors' rights generally and to general principles of equity; and (e) the execution, delivery, and performance of this Agreement do not and will not conflict with, result in a breach of, or constitute a default under any agreement, instrument, order, judgment, or decree to which such party is a party or by which it is bound.",
      ]},
      { title: "Section 6.2. Party B's Representations and Warranties", paragraphs: [
        "Party B represents and warrants to Party A that: (a) the Services shall be performed in a professional and workmanlike manner by qualified personnel with the requisite skills, experience, and expertise; (b) the Deliverables shall conform to the specifications, requirements, and acceptance criteria set forth in the applicable Exhibits; (c) the Deliverables shall be free from material defects in design, materials, and workmanship; (d) Party B has all rights, licenses, and permissions necessary to perform the Services and deliver the Deliverables; (e) the Deliverables shall not infringe, misappropriate, or otherwise violate any Intellectual Property rights of any third party; and (f) Party B shall comply with all applicable Laws in the performance of the Services.",
        "Party B further represents and warrants that none of its personnel assigned to perform the Services is subject to any restrictive covenant, non-competition agreement, or other obligation that would prevent or restrict such personnel from performing the Services or would give rise to any claim by any third party in connection with the performance of the Services.",
      ]},
      { title: "Section 6.3. Disclaimer", paragraphs: [
        "EXCEPT AS EXPRESSLY SET FORTH IN THIS AGREEMENT, NEITHER PARTY MAKES ANY REPRESENTATIONS OR WARRANTIES OF ANY KIND, WHETHER EXPRESS, IMPLIED, STATUTORY, OR OTHERWISE, AND EACH PARTY SPECIFICALLY DISCLAIMS ALL IMPLIED WARRANTIES, INCLUDING ANY IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, TITLE, AND NON-INFRINGEMENT, TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW.",
      ]},
    ],
  },
  {
    title: "ARTICLE VII - LIMITATION OF LIABILITY",
    subsections: [
      { title: "Section 7.1. Limitation on Consequential Damages", paragraphs: [
        "IN NO EVENT SHALL EITHER PARTY BE LIABLE TO THE OTHER PARTY FOR ANY INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL, PUNITIVE, OR EXEMPLARY DAMAGES, INCLUDING DAMAGES FOR LOSS OF PROFITS, LOSS OF REVENUE, LOSS OF BUSINESS OPPORTUNITIES, LOSS OF DATA, LOSS OF GOODWILL, BUSINESS INTERRUPTION, OR COST OF PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES, ARISING OUT OF OR RELATING TO THIS AGREEMENT, REGARDLESS OF THE THEORY OF LIABILITY (WHETHER IN CONTRACT, TORT, STRICT LIABILITY, OR OTHERWISE) AND REGARDLESS OF WHETHER SUCH PARTY HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.",
      ]},
      { title: "Section 7.2. Cap on Liability", paragraphs: [
        "EXCEPT WITH RESPECT TO THE OBLIGATIONS SET FORTH IN SECTIONS 7.3 AND 7.4, EACH PARTY'S TOTAL CUMULATIVE LIABILITY ARISING OUT OF OR RELATING TO THIS AGREEMENT SHALL NOT EXCEED THE AGGREGATE AMOUNT OF FEES PAID OR PAYABLE BY PARTY A TO PARTY B UNDER THIS AGREEMENT DURING THE TWELVE (12) MONTH PERIOD IMMEDIATELY PRECEDING THE EVENT GIVING RISE TO SUCH LIABILITY.",
      ]},
      { title: "Section 7.3. Exclusions from Limitation", paragraphs: [
        "The limitations of liability set forth in Sections 7.1 and 7.2 shall not apply to: (a) either party's indemnification obligations under Article VIII; (b) either party's breach of its confidentiality obligations under Article V; (c) Party B's breach of its Intellectual Property-related representations and warranties under Section 6.2(e); (d) either party's willful misconduct or gross negligence; or (e) either party's obligations to pay fees, expenses, or other amounts due under this Agreement.",
      ]},
      { title: "Section 7.4. Essential Basis", paragraphs: [
        "The parties acknowledge and agree that the limitations of liability set forth in this Article VII reflect a fair and reasonable allocation of risk between the parties, that such limitations are an essential basis of the bargain between the parties, and that the parties would not have entered into this Agreement without such limitations. The parties further acknowledge that the Fees charged by Party B reflect such allocation of risk and such limitations of liability.",
      ]},
    ],
  },
  {
    title: "ARTICLE VIII - INDEMNIFICATION",
    subsections: [
      { title: "Section 8.1. Indemnification by Party B", paragraphs: [
        "Party B shall indemnify, defend, and hold harmless Party A and its Affiliates, and their respective officers, directors, employees, agents, successors, and assigns (collectively, the 'Party A Indemnitees'), from and against any and all Claims and Damages arising out of or relating to: (a) any breach of any representation, warranty, covenant, or obligation of Party B under this Agreement; (b) any claim that the Deliverables or the Services infringe, misappropriate, or otherwise violate any Intellectual Property rights of any third party; (c) the negligence, willful misconduct, or fraud of Party B or its personnel in connection with the performance of the Services; or (d) any violation of applicable Law by Party B or its personnel in connection with the performance of the Services.",
      ]},
      { title: "Section 8.2. Indemnification by Party A", paragraphs: [
        "Party A shall indemnify, defend, and hold harmless Party B and its Affiliates, and their respective officers, directors, employees, agents, successors, and assigns (collectively, the 'Party B Indemnitees'), from and against any and all Claims and Damages arising out of or relating to: (a) any breach of any representation, warranty, covenant, or obligation of Party A under this Agreement; (b) any claim arising from Party A's use of the Deliverables in a manner not contemplated by this Agreement; (c) the negligence, willful misconduct, or fraud of Party A or its personnel; or (d) any violation of applicable Law by Party A or its personnel.",
      ]},
      { title: "Section 8.3. Indemnification Procedures", paragraphs: [
        "The Indemnified Party shall provide the Indemnifying Party with prompt written notice of any Claim for which indemnification is sought under this Article VIII, provided that the failure to provide such prompt notice shall not relieve the Indemnifying Party of its indemnification obligations except to the extent that the Indemnifying Party is materially prejudiced by such failure.",
        "The Indemnifying Party shall have the right to assume and control the defense of any such Claim, at its own cost and expense, with counsel of its choosing that is reasonably acceptable to the Indemnified Party. The Indemnified Party shall cooperate with the Indemnifying Party in the defense of such Claim and shall have the right to participate in such defense at its own cost and expense.",
        "The Indemnifying Party shall not settle or compromise any Claim without the prior written consent of the Indemnified Party, which consent shall not be unreasonably withheld, conditioned, or delayed, unless such settlement (a) involves only the payment of money damages by the Indemnifying Party, (b) does not involve any admission of liability or wrongdoing by the Indemnified Party, and (c) includes a complete and unconditional release of the Indemnified Party from all liability with respect to such Claim.",
      ]},
    ],
  },
  {
    title: "ARTICLE IX - INSURANCE",
    subsections: [
      { title: "Section 9.1. Required Coverage", paragraphs: [
        "During the Term and for a period of two (2) years following the expiration or termination of this Agreement, Party B shall maintain, at its own cost and expense, the following insurance coverage with insurance carriers rated A- VII or better by A.M. Best: (a) commercial general liability insurance with limits of not less than Two Million Dollars ($2,000,000) per occurrence and Five Million Dollars ($5,000,000) in the aggregate; (b) professional liability (errors and omissions) insurance with limits of not less than Five Million Dollars ($5,000,000) per claim and in the aggregate; (c) workers' compensation insurance as required by applicable Law; (d) employer's liability insurance with limits of not less than One Million Dollars ($1,000,000) per occurrence; and (e) cyber liability insurance with limits of not less than Five Million Dollars ($5,000,000) per claim and in the aggregate.",
        "Party B shall cause Party A to be named as an additional insured under Party B's commercial general liability and cyber liability insurance policies. Party B shall provide Party A with certificates of insurance evidencing the required coverage upon request.",
      ]},
    ],
  },
  {
    title: "ARTICLE X - TERM AND TERMINATION",
    subsections: [
      { title: "Section 10.1. Term", paragraphs: [
        "This Agreement shall commence on the Effective Date and shall continue in effect for a period of five (5) years, unless earlier terminated in accordance with this Article X (the 'Initial Term'). Thereafter, this Agreement shall automatically renew for successive one (1) year periods (each, a 'Renewal Term') unless either party provides written notice of non-renewal to the other party at least ninety (90) days prior to the expiration of the then-current Term.",
      ]},
      { title: "Section 10.2. Termination for Cause", paragraphs: [
        "Either party may terminate this Agreement immediately upon written notice to the other party if: (a) the other party commits a material breach of this Agreement and fails to cure such breach within thirty (30) days after receipt of written notice specifying the nature of the breach in reasonable detail; (b) the other party becomes insolvent, files or has filed against it a petition in bankruptcy, makes an assignment for the benefit of creditors, or has a receiver or trustee appointed for all or substantially all of its assets; or (c) the other party is dissolved, liquidated, or ceases to conduct business in the ordinary course.",
      ]},
      { title: "Section 10.3. Termination for Convenience", paragraphs: [
        "Party A may terminate this Agreement for convenience at any time upon sixty (60) days' prior written notice to Party B. In the event of such termination, Party A shall pay Party B for all Services performed and Deliverables delivered through the effective date of termination, together with any non-cancellable costs and expenses reasonably incurred by Party B in reliance on this Agreement prior to Party B's receipt of the termination notice.",
      ]},
      { title: "Section 10.4. Effects of Termination", paragraphs: [
        "Upon the expiration or termination of this Agreement for any reason: (a) all licenses granted by one party to the other under this Agreement shall immediately terminate, except as expressly provided herein; (b) each party shall promptly return or destroy all Confidential Information of the other party in accordance with Section 5.4; (c) Party B shall promptly deliver to Party A all completed and in-progress Deliverables and all Party A materials in Party B's possession or control; and (d) each party shall pay to the other party all amounts due and owing under this Agreement as of the effective date of termination.",
        "The following provisions shall survive the expiration or termination of this Agreement: Article I (Definitions), Article IV (Intellectual Property), Article V (Confidentiality), Article VI (Representations and Warranties) (for the applicable survival period), Article VII (Limitation of Liability), Article VIII (Indemnification), this Section 10.4, and Article XI (General Provisions).",
      ]},
    ],
  },
  {
    title: "ARTICLE XI - GENERAL PROVISIONS",
    subsections: [
      { title: "Section 11.1. Governing Law", paragraphs: [
        "This Agreement shall be governed by and construed in accordance with the laws of the State of New York, without giving effect to any choice or conflict of law provision or rule that would cause the application of the laws of any jurisdiction other than the State of New York. The United Nations Convention on Contracts for the International Sale of Goods shall not apply to this Agreement.",
      ]},
      { title: "Section 11.2. Dispute Resolution", paragraphs: [
        "Any dispute, controversy, or claim arising out of or relating to this Agreement, or the breach, termination, or invalidity thereof, shall first be submitted to mediation in accordance with the mediation rules of the American Arbitration Association then in effect. The mediation shall take place in New York, New York, and shall be conducted by a single mediator mutually agreed upon by the parties.",
        "If the dispute is not resolved through mediation within sixty (60) days after the commencement of mediation, either party may submit the dispute to binding arbitration administered by the American Arbitration Association in accordance with its Commercial Arbitration Rules. The arbitration shall take place in New York, New York, and shall be conducted by a panel of three (3) arbitrators. The arbitrators shall have expertise in commercial transactions and technology law.",
        "The arbitrators shall issue a written, reasoned decision, and judgment upon the award rendered by the arbitrators may be entered in any court having jurisdiction thereof. The arbitrators shall have the authority to award any remedy or relief that a court of competent jurisdiction could order, including specific performance, injunctive relief, and reasonable attorneys' fees and costs. The parties agree that the arbitration proceedings and any related discovery shall be conducted on a confidential basis.",
      ]},
      { title: "Section 11.3. Notices", paragraphs: [
        "All notices, requests, demands, consents, and other communications required or permitted under this Agreement shall be in writing and shall be deemed to have been duly given: (a) when delivered personally; (b) when sent by confirmed electronic mail (with a copy sent by first class mail); (c) one (1) Business Day after deposit with a nationally recognized overnight courier service, with all charges prepaid; or (d) three (3) Business Days after being sent by certified or registered mail, return receipt requested, postage prepaid, to the addresses set forth on the signature page hereof or to such other address as either party may designate by notice to the other party in accordance with this Section.",
      ]},
      { title: "Section 11.4. Assignment", paragraphs: [
        "Neither party may assign or transfer this Agreement, in whole or in part, by operation of law or otherwise, without the prior written consent of the other party, which consent shall not be unreasonably withheld, conditioned, or delayed. Notwithstanding the foregoing, either party may assign this Agreement without the other party's consent to an Affiliate or in connection with a Change of Control, provided that the assignee assumes all of the assignor's obligations under this Agreement. Any attempted assignment in violation of this Section shall be null and void. Subject to the foregoing, this Agreement shall be binding upon and inure to the benefit of the parties and their respective successors and permitted assigns.",
      ]},
      { title: "Section 11.5. Force Majeure", paragraphs: [
        "Neither party shall be liable for any failure or delay in performing its obligations under this Agreement (other than payment obligations) to the extent that such failure or delay results from circumstances beyond the reasonable control of that party, including acts of God, natural disasters, epidemics, pandemics, war, terrorism, riots, civil unrest, government actions, embargoes, sanctions, labor disputes, strikes, fire, flood, earthquake, hurricane, power outages, telecommunications failures, or cyberattacks (each, a 'Force Majeure Event').",
        "The affected party shall provide prompt written notice to the other party of the Force Majeure Event and its expected duration, and shall use commercially reasonable efforts to mitigate the effects of the Force Majeure Event and resume performance as soon as practicable. If a Force Majeure Event continues for more than ninety (90) consecutive days, either party may terminate this Agreement upon thirty (30) days' written notice to the other party.",
      ]},
      { title: "Section 11.6. Independent Contractors", paragraphs: [
        "The relationship between the parties is that of independent contractors. Nothing in this Agreement shall be construed to create a partnership, joint venture, agency, employment, or franchise relationship between the parties. Neither party shall have the authority to bind the other party or to incur any obligation on behalf of the other party without the other party's prior written consent.",
        "Party B's personnel are not employees of Party A and shall not be entitled to any employee benefits from Party A, including health insurance, retirement benefits, paid time off, or workers' compensation. Party B shall be solely responsible for the payment of all compensation, benefits, and taxes for its personnel.",
      ]},
      { title: "Section 11.7. Non-Solicitation", paragraphs: [
        "During the Term and for a period of one (1) year following the expiration or termination of this Agreement, neither party shall, directly or indirectly, solicit or recruit, or attempt to solicit or recruit, any employee or contractor of the other party who was involved in the performance of the Services, without the prior written consent of the other party. This restriction shall not apply to (a) general solicitations of employment not specifically directed at the other party's employees or contractors, or (b) any individual who has ceased to be employed by or engaged by the other party for a period of at least six (6) months.",
      ]},
      { title: "Section 11.8. Entire Agreement", paragraphs: [
        "This Agreement, together with all Exhibits and Schedules attached hereto and incorporated herein by reference, constitutes the entire agreement between the parties with respect to the subject matter hereof and supersedes all prior and contemporaneous agreements, understandings, negotiations, and discussions, whether written or oral, between the parties with respect to such subject matter. Each party acknowledges that it has not relied on any statement, representation, warranty, or agreement of the other party except for those expressly set forth in this Agreement.",
      ]},
      { title: "Section 11.9. Amendments and Waivers", paragraphs: [
        "No amendment, modification, supplement, or waiver of any provision of this Agreement shall be effective unless in writing and signed by authorized representatives of both parties. No waiver of any provision of this Agreement shall constitute a waiver of any other provision or of the same provision on any other occasion. No failure or delay by either party in exercising any right or remedy under this Agreement shall operate as a waiver of such right or remedy.",
      ]},
      { title: "Section 11.10. Severability", paragraphs: [
        "If any provision of this Agreement is held to be invalid, illegal, or unenforceable by a court of competent jurisdiction, the validity, legality, and enforceability of the remaining provisions shall not be affected or impaired thereby. The parties shall negotiate in good faith to replace any invalid, illegal, or unenforceable provision with a valid, legal, and enforceable provision that achieves, to the greatest extent possible, the economic, business, and other purposes of the invalid, illegal, or unenforceable provision.",
      ]},
      { title: "Section 11.11. Counterparts", paragraphs: [
        "This Agreement may be executed in one or more counterparts, each of which shall be deemed an original, but all of which together shall constitute one and the same instrument. Signatures transmitted by facsimile or electronic means (including PDF) shall be deemed original signatures for all purposes.",
      ]},
      { title: "Section 11.12. Third-Party Beneficiaries", paragraphs: [
        "Except as expressly set forth in Article VIII with respect to the Party A Indemnitees and the Party B Indemnitees, this Agreement is for the sole benefit of the parties and their respective successors and permitted assigns, and nothing in this Agreement, express or implied, is intended to or shall confer upon any other Person any legal or equitable right, benefit, or remedy of any nature whatsoever under or by reason of this Agreement.",
      ]},
    ],
  },
];

// Generate additional schedule/exhibit content to reach ~100 pages
const exhibits = [
  {
    title: "EXHIBIT A - STATEMENT OF WORK",
    content: [
      "1. PROJECT OVERVIEW",
      "This Statement of Work describes the professional services to be provided by Party B to Party A in connection with the development, implementation, and deployment of a comprehensive enterprise resource planning (ERP) system (the 'Project'). The Project encompasses the design, development, testing, deployment, and post-deployment support of an integrated software platform that will serve as the primary operational backbone for Party A's business operations across all departments and geographic locations within the Territory.",
      "The ERP system shall include, without limitation, the following functional modules: (a) Financial Management, including general ledger, accounts payable, accounts receivable, fixed assets, cash management, and financial reporting; (b) Human Capital Management, including employee records, payroll processing, benefits administration, time and attendance, talent acquisition, and performance management; (c) Supply Chain Management, including procurement, inventory management, warehouse management, order management, and logistics; (d) Customer Relationship Management, including sales force automation, marketing automation, customer service, and analytics; and (e) Business Intelligence and Analytics, including real-time dashboards, ad hoc reporting, data visualization, and predictive analytics.",
      "2. PROJECT PHASES",
      "Phase 1 - Discovery and Planning (Months 1-3): During this phase, Party B shall conduct a comprehensive assessment of Party A's existing business processes, systems, and data. Party B shall interview key stakeholders across all departments to understand current workflows, pain points, and requirements. Party B shall deliver a detailed requirements document, a gap analysis comparing current state to desired future state, a preliminary system architecture design, and a comprehensive project plan with milestones and timelines.",
      "Phase 2 - Design and Architecture (Months 4-6): Based on the approved requirements document, Party B shall develop a detailed system design, including data models, integration architecture, user interface wireframes, security architecture, and technical specifications for all functional modules. Party B shall conduct design review sessions with Party A's stakeholders and incorporate feedback into the final design documents.",
      "Phase 3 - Development and Configuration (Months 7-14): Party B shall develop, configure, and customize the ERP system in accordance with the approved design documents. Development shall follow an agile methodology with two-week sprints and regular demonstrations to Party A's project team. Party B shall maintain a development environment, a testing environment, and a staging environment throughout this phase.",
      "Phase 4 - Testing and Quality Assurance (Months 15-17): Party B shall conduct comprehensive testing of the ERP system, including unit testing, integration testing, system testing, performance testing, security testing, and user acceptance testing. Party B shall develop and maintain a detailed test plan and test cases covering all functional requirements and non-functional requirements. Party A shall participate in user acceptance testing and shall have the right to identify defects and request corrections.",
      "Phase 5 - Deployment and Go-Live (Month 18): Party B shall deploy the ERP system to Party A's production environment in accordance with the deployment plan approved by the Steering Committee. Party B shall provide on-site support during the go-live period, including a team of at least ten (10) qualified support personnel available twenty-four (24) hours per day, seven (7) days per week, for a period of four (4) weeks following the go-live date.",
      "Phase 6 - Post-Deployment Support (Months 19-24): Following successful deployment, Party B shall provide ongoing support services, including bug fixes, system maintenance, performance optimization, and minor enhancements. Party B shall maintain a service desk staffed by qualified support personnel during Party A's normal business hours. Party B shall respond to critical issues within one (1) hour, high-priority issues within four (4) hours, and routine issues within one (1) Business Day.",
      "3. STAFFING REQUIREMENTS",
      "Party B shall assign the following key personnel to the Project: (a) a Project Director with at least fifteen (15) years of experience managing large-scale ERP implementations; (b) a Technical Architect with at least ten (10) years of experience in enterprise software architecture; (c) a Lead Developer with at least eight (8) years of experience in full-stack software development; (d) a Quality Assurance Lead with at least eight (8) years of experience in software testing and quality assurance; (e) a Data Migration Specialist with at least five (5) years of experience in data migration and ETL processes; and (f) a Change Management Lead with at least seven (7) years of experience in organizational change management.",
      "Party B shall not remove or replace any key personnel without the prior written consent of Party A, unless such removal is due to the individual's resignation, termination for cause, or inability to work due to illness or disability. In the event of any such removal, Party B shall promptly propose a replacement with substantially equivalent qualifications and experience for Party A's approval.",
    ],
  },
  {
    title: "EXHIBIT B - DELIVERABLES AND ACCEPTANCE CRITERIA",
    content: [
      "1. DELIVERABLES SCHEDULE",
      "The following table sets forth the Deliverables to be provided by Party B, together with the delivery dates and acceptance criteria for each Deliverable:",
      "Deliverable 1.1 - Requirements Document: A comprehensive document detailing all functional and non-functional requirements for the ERP system, organized by functional module. Due Date: End of Month 3. Acceptance Criteria: The document must cover all in-scope functional modules, address all requirements identified during the discovery phase, and be approved by the Steering Committee.",
      "Deliverable 1.2 - Gap Analysis Report: A detailed analysis of the gaps between Party A's current systems and processes and the desired future state, including recommendations for addressing each gap. Due Date: End of Month 3. Acceptance Criteria: The report must identify all material gaps, provide actionable recommendations, and include estimated effort and cost for each recommendation.",
      "Deliverable 2.1 - System Architecture Document: A comprehensive document describing the technical architecture of the ERP system, including system components, integration points, data flows, security architecture, infrastructure requirements, and scalability considerations. Due Date: End of Month 5. Acceptance Criteria: The architecture must comply with Party A's technology standards, support the required scalability and performance requirements, and be approved by Party A's Chief Technology Officer.",
      "Deliverable 2.2 - User Interface Design: Detailed wireframes and visual designs for all user-facing screens and interfaces of the ERP system, including responsive designs for desktop, tablet, and mobile devices. Due Date: End of Month 6. Acceptance Criteria: Designs must comply with Party A's brand guidelines, meet WCAG 2.1 Level AA accessibility standards, and be approved by Party A's User Experience team.",
      "Deliverable 3.1 - Functional Modules: Fully developed and configured functional modules for Financial Management, Human Capital Management, Supply Chain Management, Customer Relationship Management, and Business Intelligence and Analytics. Due Date: End of Month 14. Acceptance Criteria: Each module must pass all defined unit tests and integration tests, meet the performance requirements specified in the non-functional requirements document, and demonstrate all required functionality in sprint review demonstrations.",
      "Deliverable 4.1 - Test Reports: Comprehensive test reports documenting the results of all testing phases, including test coverage metrics, defect logs, and resolution status. Due Date: End of Month 17. Acceptance Criteria: Test coverage must meet or exceed ninety-five percent (95%) of all documented requirements, all critical and high-priority defects must be resolved, and the system must meet all performance benchmarks.",
      "Deliverable 5.1 - Production Deployment: Successful deployment of the ERP system to Party A's production environment, including data migration from legacy systems, configuration of production infrastructure, and completion of go-live checklist. Due Date: End of Month 18. Acceptance Criteria: The system must be fully operational in the production environment, all migrated data must be validated for accuracy and completeness, and the system must meet all availability and performance requirements.",
      "Deliverable 6.1 - Training Materials: Comprehensive training materials for all user roles, including user guides, quick reference cards, training videos, and online help documentation. Due Date: End of Month 17. Acceptance Criteria: Materials must cover all system functionality relevant to each user role, be written in clear and accessible language, and be available in both English and Spanish.",
      "2. ACCEPTANCE PROCESS",
      "Upon delivery of each Deliverable, Party A shall have thirty (30) Business Days to review and test the Deliverable against the applicable acceptance criteria (the 'Review Period'). During the Review Period, Party A shall provide Party B with written notice of acceptance or rejection of the Deliverable.",
      "If Party A rejects a Deliverable, Party A shall provide a detailed description of the deficiencies and the specific acceptance criteria that have not been met. Party B shall correct the identified deficiencies and redeliver the Deliverable within fifteen (15) Business Days. Party A shall then have an additional fifteen (15) Business Day Review Period to review the corrected Deliverable.",
      "If a Deliverable is rejected a second time, the parties shall escalate the matter to the Steering Committee for resolution. If the Steering Committee is unable to resolve the issue within thirty (30) days, either party may exercise its termination rights under Article X of the Agreement.",
    ],
  },
  {
    title: "EXHIBIT C - FEE SCHEDULE",
    content: [
      "1. PROJECT FEES",
      "The total fixed fee for the Project shall be Twelve Million Five Hundred Thousand Dollars ($12,500,000.00), payable in accordance with the milestone payment schedule set forth below. The fixed fee includes all labor costs, overhead, profit, and other costs and expenses incurred by Party B in connection with the performance of the Services, except for reimbursable expenses as provided in Section 3.3 of the Agreement.",
      "Milestone 1 - Project Initiation: Five percent (5%) of the total fixed fee ($625,000.00), payable upon execution of this Agreement.",
      "Milestone 2 - Completion of Phase 1 (Discovery and Planning): Ten percent (10%) of the total fixed fee ($1,250,000.00), payable upon Party A's acceptance of Deliverables 1.1 and 1.2.",
      "Milestone 3 - Completion of Phase 2 (Design and Architecture): Fifteen percent (15%) of the total fixed fee ($1,875,000.00), payable upon Party A's acceptance of Deliverables 2.1 and 2.2.",
      "Milestone 4 - Completion of Phase 3 (Development and Configuration): Thirty-five percent (35%) of the total fixed fee ($4,375,000.00), payable in seven (7) equal monthly installments of $625,000.00 each during the development phase, contingent upon satisfactory progress as determined by the Steering Committee.",
      "Milestone 5 - Completion of Phase 4 (Testing and Quality Assurance): Fifteen percent (15%) of the total fixed fee ($1,875,000.00), payable upon Party A's acceptance of Deliverable 4.1.",
      "Milestone 6 - Successful Go-Live: Ten percent (10%) of the total fixed fee ($1,250,000.00), payable upon Party A's acceptance of Deliverable 5.1.",
      "Milestone 7 - Completion of Post-Deployment Support: Ten percent (10%) of the total fixed fee ($1,250,000.00), payable upon successful completion of the post-deployment support period.",
      "2. HOURLY RATES FOR CHANGE ORDERS",
      "Services performed pursuant to approved change orders shall be billed at the following hourly rates: Project Director: $450 per hour; Technical Architect: $400 per hour; Senior Developer: $350 per hour; Developer: $275 per hour; Quality Assurance Analyst: $250 per hour; Business Analyst: $275 per hour; Project Manager: $300 per hour; Data Migration Specialist: $300 per hour; Change Management Consultant: $325 per hour; Training Specialist: $225 per hour.",
      "The foregoing hourly rates shall be fixed for the Initial Term and shall be subject to an annual increase not to exceed three percent (3%) during any Renewal Term.",
      "3. ANNUAL MAINTENANCE AND SUPPORT FEES",
      "Following the completion of the post-deployment support period described in Phase 6, Party B shall provide ongoing maintenance and support services for an annual fee of One Million Two Hundred Fifty Thousand Dollars ($1,250,000.00) per year, payable in equal quarterly installments. The annual maintenance and support fee includes: (a) bug fixes and patches; (b) minor enhancements (not to exceed one hundred (100) hours of development effort per quarter); (c) system monitoring and performance optimization; (d) service desk support during Party A's normal business hours; and (e) two (2) major version upgrades per year.",
    ],
  },
  {
    title: "EXHIBIT D - SERVICE LEVEL AGREEMENT",
    content: [
      "1. AVAILABILITY",
      "Party B shall ensure that the ERP system is available for use by Party A's authorized users at least ninety-nine point nine percent (99.9%) of the time during each calendar month, measured on a twenty-four (24) hours per day, seven (7) days per week basis, excluding scheduled maintenance windows ('Uptime Requirement'). Scheduled maintenance windows shall be limited to no more than four (4) hours per week and shall be scheduled during non-business hours with at least forty-eight (48) hours' prior notice to Party A.",
      "In the event that Party B fails to meet the Uptime Requirement in any calendar month, Party A shall be entitled to service credits as follows: (a) for availability between 99.0% and 99.9%, a credit of five percent (5%) of the monthly maintenance and support fee; (b) for availability between 98.0% and 99.0%, a credit of ten percent (10%) of the monthly maintenance and support fee; (c) for availability between 95.0% and 98.0%, a credit of twenty percent (20%) of the monthly maintenance and support fee; and (d) for availability below 95.0%, a credit of thirty percent (30%) of the monthly maintenance and support fee.",
      "2. RESPONSE AND RESOLUTION TIMES",
      "Party B shall respond to and resolve incidents in accordance with the following service levels based on incident severity: Critical (Severity 1) - System is completely unavailable or a core business function is inoperable: Response Time: fifteen (15) minutes; Resolution Time: four (4) hours. Party B shall assign its most qualified personnel and work continuously until the issue is resolved.",
      "High (Severity 2) - A major function is significantly impaired but the system remains operational: Response Time: one (1) hour; Resolution Time: eight (8) business hours. Party B shall assign senior technical personnel and provide regular status updates at least every two (2) hours.",
      "Medium (Severity 3) - A non-critical function is impaired or a workaround is available: Response Time: four (4) business hours; Resolution Time: three (3) Business Days. Party B shall provide a root cause analysis and permanent fix within the resolution timeframe.",
      "Low (Severity 4) - A minor issue or enhancement request that does not materially affect system functionality: Response Time: one (1) Business Day; Resolution Time: ten (10) Business Days. Party B shall schedule the fix or enhancement in the next available maintenance window or release cycle.",
      "3. PERFORMANCE REQUIREMENTS",
      "The ERP system shall meet the following performance requirements at all times during normal operating conditions: (a) average page load time shall not exceed two (2) seconds for ninety-five percent (95%) of all page requests; (b) batch processing jobs shall complete within the time windows specified in the system documentation; (c) the system shall support at least five hundred (500) concurrent users without degradation in performance; (d) database queries shall return results within three (3) seconds for ninety-nine percent (99%) of all queries; and (e) API response times shall not exceed five hundred (500) milliseconds for ninety-five percent (95%) of all API calls.",
      "4. REPORTING",
      "Party B shall provide Party A with monthly service level reports within five (5) Business Days following the end of each calendar month. Each report shall include: (a) actual system availability for the month compared to the Uptime Requirement; (b) a summary of all incidents by severity, including response times, resolution times, and root cause analyses; (c) performance metrics compared to the performance requirements; (d) a summary of all maintenance activities performed during the month; (e) a forecast of planned maintenance activities for the upcoming month; and (f) any service credits earned by Party A during the month.",
      "5. CONTINUOUS IMPROVEMENT",
      "Party B shall implement a continuous improvement program to identify and implement opportunities to improve system performance, reliability, and efficiency. Party B shall conduct quarterly service level reviews with Party A to discuss service level performance, identify trends and areas for improvement, and agree on action items. Party B shall maintain a knowledge base of known issues and resolutions and shall implement proactive monitoring to detect and resolve potential issues before they impact system availability or performance.",
    ],
  },
  {
    title: "EXHIBIT E - DATA SECURITY AND PRIVACY REQUIREMENTS",
    content: [
      "1. GENERAL SECURITY REQUIREMENTS",
      "Party B shall implement and maintain a comprehensive information security program that includes administrative, technical, and physical safeguards designed to protect the confidentiality, integrity, and availability of Party A's data and systems. Party B's information security program shall comply with industry best practices, including the ISO/IEC 27001 standard, the NIST Cybersecurity Framework, and the CIS Controls.",
      "Party B shall maintain current SOC 2 Type II certification covering security, availability, processing integrity, confidentiality, and privacy trust services criteria. Party B shall provide Party A with copies of its SOC 2 reports upon request and shall promptly notify Party A of any material findings or exceptions identified in such reports.",
      "2. ACCESS CONTROLS",
      "Party B shall implement role-based access controls to ensure that access to Party A's data and systems is limited to authorized personnel on a need-to-know basis. Party B shall implement multi-factor authentication for all administrative access and for all remote access to Party A's systems. Party B shall maintain detailed access logs and shall review access rights on at least a quarterly basis to ensure that access is appropriate and current.",
      "All Party B personnel with access to Party A's data or systems shall undergo background checks, sign confidentiality agreements, and complete security awareness training before being granted access. Party B shall promptly revoke access for any personnel who no longer require access or who have been terminated or reassigned.",
      "3. ENCRYPTION",
      "Party B shall encrypt all Party A data at rest using AES-256 encryption or equivalent, and all Party A data in transit using TLS 1.2 or higher. Party B shall implement key management procedures that comply with industry best practices, including secure key generation, storage, rotation, and destruction.",
      "4. INCIDENT RESPONSE",
      "Party B shall maintain an incident response plan that includes procedures for detecting, reporting, investigating, containing, and remediating security incidents. Party B shall notify Party A of any security incident involving Party A's data within twenty-four (24) hours of detection. Party B shall cooperate with Party A in investigating and remediating any security incident and shall provide Party A with a detailed incident report within five (5) Business Days following the resolution of the incident.",
      "In the event of a data breach involving Party A's personal data, Party B shall (a) notify Party A within the timeframes required by applicable Law (but in no event later than forty-eight (48) hours after detection), (b) provide Party A with all information necessary to comply with applicable breach notification requirements, (c) cooperate with Party A in notifying affected individuals and regulatory authorities, and (d) take all necessary steps to remediate the breach and prevent future occurrences.",
      "5. DATA PRIVACY",
      "Party B shall process Party A's personal data only in accordance with Party A's documented instructions and applicable privacy laws, including the General Data Protection Regulation (GDPR), the California Consumer Privacy Act (CCPA), and any other applicable data protection laws. Party B shall not process Party A's personal data for any purpose other than providing the Services, and shall not sell, share, or otherwise disclose Party A's personal data to any third party except as expressly authorized by Party A or required by applicable Law.",
      "Party B shall implement appropriate technical and organizational measures to ensure a level of security appropriate to the risk, taking into account the state of the art, the costs of implementation, the nature, scope, context, and purposes of processing, and the risks to the rights and freedoms of data subjects. Party B shall assist Party A in responding to data subject requests, conducting data protection impact assessments, and complying with its obligations under applicable privacy laws.",
      "6. AUDIT RIGHTS",
      "Party A shall have the right, upon thirty (30) days' prior written notice, to audit Party B's compliance with the security and privacy requirements set forth in this Exhibit, including the right to inspect Party B's facilities, systems, and records, and to interview Party B's personnel. Party B shall cooperate fully with any such audit and shall promptly remediate any deficiencies identified during the audit. Party A may exercise its audit rights no more than once per calendar year, unless a security incident or material deficiency has been identified, in which case Party A may conduct additional audits as reasonably necessary.",
    ],
  },
];

// Build the document
let bodyContent = "";

// Title page
bodyContent += wParagraph("MASTER SERVICES AGREEMENT", { bold: true, size: 36, spacing: { before: 4000, after: 200 } });
bodyContent += wParagraph("", { spacing: { before: 200, after: 200 } });
bodyContent += wParagraph("by and between", { size: 24, spacing: { before: 200, after: 200 } });
bodyContent += wParagraph("", { spacing: { before: 200, after: 200 } });
bodyContent += wParagraph("ACME GLOBAL ENTERPRISES, INC.", { bold: true, size: 28, spacing: { before: 200, after: 200 } });
bodyContent += wParagraph("('Party A')", { size: 24, spacing: { before: 100, after: 400 } });
bodyContent += wParagraph("and", { size: 24, spacing: { before: 200, after: 200 } });
bodyContent += wParagraph("", { spacing: { before: 200, after: 200 } });
bodyContent += wParagraph("PINNACLE TECHNOLOGY SOLUTIONS, LLC", { bold: true, size: 28, spacing: { before: 200, after: 200 } });
bodyContent += wParagraph("('Party B')", { size: 24, spacing: { before: 100, after: 400 } });
bodyContent += wParagraph("", { spacing: { before: 200, after: 200 } });
bodyContent += wParagraph("Effective Date: January 1, 2026", { size: 24, spacing: { before: 400, after: 200 } });
bodyContent += pageBreak();

// Table of contents placeholder
bodyContent += wParagraph("TABLE OF CONTENTS", { bold: true, size: 28, spacing: { before: 400, after: 400 } });
for (const section of sections) {
  bodyContent += wParagraph(section.title, { size: 22, spacing: { before: 100, after: 100 } });
}
for (const exhibit of exhibits) {
  bodyContent += wParagraph(exhibit.title, { size: 22, spacing: { before: 100, after: 100 } });
}
bodyContent += pageBreak();

// Main body sections
for (const section of sections) {
  bodyContent += wParagraph(section.title, { bold: true, size: 28, spacing: { before: 400, after: 200 }, keepNext: true });

  for (const sub of section.subsections) {
    if (sub.title) {
      bodyContent += wParagraph(sub.title, { bold: true, size: 24, spacing: { before: 300, after: 100 }, keepNext: true });
    }
    for (const para of sub.paragraphs) {
      bodyContent += wParagraph(para, { size: 22, spacing: { before: 100, after: 100 } });
    }
  }
}

bodyContent += pageBreak();

// Signature page
bodyContent += wParagraph("IN WITNESS WHEREOF, the parties have executed this Agreement as of the Effective Date.", { size: 22, spacing: { before: 400, after: 600 } });

bodyContent += wParagraph("ACME GLOBAL ENTERPRISES, INC.", { bold: true, size: 22, spacing: { before: 400, after: 200 } });
bodyContent += wParagraph("By: ___________________________________", { size: 22, spacing: { before: 200, after: 100 } });
bodyContent += wParagraph("Name: John A. Richardson", { size: 22, spacing: { before: 100, after: 100 } });
bodyContent += wParagraph("Title: Chief Executive Officer", { size: 22, spacing: { before: 100, after: 100 } });
bodyContent += wParagraph("Date: ___________________________________", { size: 22, spacing: { before: 100, after: 400 } });

bodyContent += wParagraph("PINNACLE TECHNOLOGY SOLUTIONS, LLC", { bold: true, size: 22, spacing: { before: 400, after: 200 } });
bodyContent += wParagraph("By: ___________________________________", { size: 22, spacing: { before: 200, after: 100 } });
bodyContent += wParagraph("Name: Sarah M. Chen", { size: 22, spacing: { before: 100, after: 100 } });
bodyContent += wParagraph("Title: Managing Director", { size: 22, spacing: { before: 100, after: 100 } });
bodyContent += wParagraph("Date: ___________________________________", { size: 22, spacing: { before: 100, after: 400 } });

bodyContent += pageBreak();

// Exhibits
for (const exhibit of exhibits) {
  bodyContent += wParagraph(exhibit.title, { bold: true, size: 28, spacing: { before: 400, after: 300 }, keepNext: true });
  for (const para of exhibit.content) {
    const isHeading = /^\d+\.\s+[A-Z]/.test(para);
    if (isHeading) {
      bodyContent += wParagraph(para, { bold: true, size: 24, spacing: { before: 300, after: 100 }, keepNext: true });
    } else {
      bodyContent += wParagraph(para, { size: 22, spacing: { before: 100, after: 100 } });
    }
  }
  bodyContent += pageBreak();
}

// Add additional boilerplate schedules to pad to ~100 pages
const scheduleTopics = [
  { title: "SCHEDULE 1 - APPROVED SUBCONTRACTORS", items: [
    "The following subcontractors have been approved by Party A for use by Party B in connection with the performance of the Services. Party B shall not engage any subcontractor not listed herein without the prior written approval of Party A.",
    "1. DataMigrate Solutions, Inc. - Scope: Data migration services, ETL development, and data quality management. Location: Austin, Texas. Key Personnel: James Morrison, Senior Data Architect. Party B shall remain fully responsible for all work performed by DataMigrate Solutions, Inc. and shall ensure that DataMigrate Solutions, Inc. complies with all terms and conditions of this Agreement, including the confidentiality and data security provisions.",
    "2. CloudOps Partners, LLC - Scope: Cloud infrastructure management, DevOps automation, and production environment monitoring. Location: Seattle, Washington. Key Personnel: Maria Santos, Cloud Operations Manager. CloudOps Partners, LLC shall maintain all certifications required by the applicable cloud service providers and shall comply with Party A's cloud security policies.",
    "3. UX Design Studio, Ltd. - Scope: User interface design, user experience research, and accessibility compliance testing. Location: New York, New York. Key Personnel: David Park, Creative Director. All design deliverables produced by UX Design Studio, Ltd. shall comply with Party A's brand guidelines and WCAG 2.1 Level AA accessibility standards.",
    "4. QualityFirst Testing Services, Inc. - Scope: Automated testing, performance testing, and security penetration testing. Location: Chicago, Illinois. Key Personnel: Lisa Thompson, QA Director. QualityFirst Testing Services, Inc. shall use industry-standard testing tools and methodologies and shall provide detailed test reports in accordance with the requirements set forth in Exhibit B.",
    "5. Compliance Advisors Group, LLP - Scope: Regulatory compliance consulting, privacy impact assessments, and audit support. Location: Washington, D.C. Key Personnel: Robert Williams, Partner. Compliance Advisors Group, LLP shall maintain current knowledge of all applicable laws and regulations and shall advise Party B on compliance requirements throughout the Project.",
  ]},
  { title: "SCHEDULE 2 - KEY PERSONNEL", items: [
    "The following individuals have been designated as Key Personnel for the Project. Party B shall not remove or replace any Key Personnel without the prior written consent of Party A, except as provided in Exhibit A.",
    "1. Michael Johnson - Role: Project Director. Experience: 18 years of experience in enterprise software implementation. Certifications: PMP, ITIL Expert, SAFe Agilist. Allocation: 100% dedicated to the Project. Michael will be responsible for overall project governance, executive stakeholder management, and ensuring alignment with Party A's strategic objectives.",
    "2. Dr. Emily Chen - Role: Technical Architect. Experience: 15 years of experience in enterprise software architecture and distributed systems. Certifications: AWS Solutions Architect Professional, Azure Solutions Architect Expert, TOGAF Certified. Allocation: 100% dedicated to the Project. Emily will be responsible for the overall system architecture, technology stack selection, and ensuring compliance with Party A's technical standards.",
    "3. Alexander Rodriguez - Role: Lead Developer. Experience: 12 years of experience in full-stack development. Certifications: Oracle Certified Professional, Microsoft Certified: Azure Developer Associate. Allocation: 100% dedicated to the Project. Alexander will lead the development team and be responsible for code quality, development standards, and technical mentoring.",
    "4. Jennifer Williams - Role: Quality Assurance Lead. Experience: 10 years of experience in software quality assurance. Certifications: ISTQB Advanced Level Test Manager, Certified Agile Tester. Allocation: 100% dedicated to the Project. Jennifer will be responsible for the overall testing strategy, test automation framework, and quality metrics reporting.",
    "5. David Kim - Role: Data Migration Specialist. Experience: 8 years of experience in data migration and ETL development. Certifications: AWS Certified Data Analytics, Informatica PowerCenter Developer. Allocation: 100% dedicated to the Project during Phases 3-5. David will be responsible for the data migration strategy, data mapping, transformation rules, and data validation procedures.",
    "6. Patricia Moore - Role: Change Management Lead. Experience: 12 years of experience in organizational change management. Certifications: Prosci Certified Change Practitioner, Certified Professional in Learning and Performance. Allocation: 50% during Phases 1-3, 100% during Phases 4-6. Patricia will be responsible for change readiness assessments, stakeholder engagement, communication planning, and end-user training.",
  ]},
  { title: "SCHEDULE 3 - TECHNICAL ENVIRONMENT SPECIFICATIONS", items: [
    "1. PRODUCTION ENVIRONMENT",
    "The production environment shall be hosted on Amazon Web Services (AWS) in the US-East-1 (Virginia) region with disaster recovery capabilities in US-West-2 (Oregon). The production environment shall include the following components:",
    "Application Tier: A minimum of eight (8) application server instances running on AWS EC2 m6i.2xlarge instances (8 vCPUs, 32 GB RAM), deployed across multiple Availability Zones for high availability. Auto-scaling shall be configured to add additional instances when CPU utilization exceeds seventy percent (70%) for more than five (5) consecutive minutes, up to a maximum of twenty (20) instances.",
    "Database Tier: Amazon RDS for PostgreSQL 15 or later, deployed in a Multi-AZ configuration with automated failover. Primary instance: db.r6g.4xlarge (16 vCPUs, 128 GB RAM) with 2 TB of provisioned IOPS SSD storage. Read replicas: Two (2) db.r6g.2xlarge instances for read-heavy workloads. Automated backups shall be retained for thirty (30) days, with point-in-time recovery capability.",
    "Caching Layer: Amazon ElastiCache for Redis, deployed in a cluster configuration with three (3) cache.r6g.xlarge nodes across multiple Availability Zones. The caching layer shall be used for session management, frequently accessed data, and API response caching.",
    "Storage: Amazon S3 for document storage and file attachments, with server-side encryption using AWS KMS-managed keys. S3 Intelligent-Tiering shall be enabled to optimize storage costs. Cross-region replication shall be configured to the disaster recovery region.",
    "Networking: Amazon VPC with public and private subnets across three (3) Availability Zones. AWS WAF shall be deployed in front of the Application Load Balancer to protect against common web exploits. AWS Shield Advanced shall be enabled for DDoS protection. VPN connectivity shall be established between Party A's corporate network and the AWS VPC.",
    "2. STAGING ENVIRONMENT",
    "The staging environment shall mirror the production environment architecture at reduced scale: four (4) application server instances (m6i.xlarge), one (1) database instance (db.r6g.2xlarge) with Multi-AZ, and one (1) ElastiCache node (cache.r6g.large). The staging environment shall be used for pre-production testing, performance validation, and user acceptance testing.",
    "3. DEVELOPMENT ENVIRONMENT",
    "The development environment shall include: two (2) application server instances (m6i.large), one (1) database instance (db.r6g.xlarge) without Multi-AZ, and one (1) ElastiCache node (cache.r6g.medium). Each developer shall have a local development environment with Docker containers replicating the application stack.",
    "4. CI/CD PIPELINE",
    "Party B shall implement a continuous integration and continuous deployment pipeline using AWS CodePipeline, CodeBuild, and CodeDeploy. The pipeline shall include automated code compilation, unit test execution, static code analysis (using SonarQube), security scanning (using Snyk), integration test execution, and automated deployment to the development and staging environments. Deployment to production shall require manual approval by authorized personnel from both parties.",
    "5. MONITORING AND OBSERVABILITY",
    "Party B shall implement comprehensive monitoring and observability using the following tools: Amazon CloudWatch for infrastructure monitoring, application performance monitoring, and log aggregation; AWS X-Ray for distributed tracing; PagerDuty for incident alerting and on-call management; Grafana for custom dashboards and visualization; and ELK Stack (Elasticsearch, Logstash, Kibana) for centralized log management and analysis. Alerts shall be configured for all critical metrics, including system availability, response times, error rates, and resource utilization.",
  ]},
  { title: "SCHEDULE 4 - DATA MIGRATION PLAN", items: [
    "1. OVERVIEW",
    "This Data Migration Plan describes the approach, methodology, and procedures for migrating data from Party A's legacy systems to the new ERP system. The data migration effort encompasses the extraction, transformation, cleansing, validation, and loading of data from the following source systems: (a) Legacy ERP System (SAP R/3), (b) Human Resources Information System (Workday), (c) Customer Relationship Management System (Salesforce), (d) Financial Reporting System (Oracle Hyperion), and (e) various departmental databases and spreadsheets.",
    "2. MIGRATION APPROACH",
    "The data migration shall follow a phased approach, with each phase corresponding to a functional module of the new ERP system. The phases shall be executed in the following order: (1) Master Data (customers, vendors, employees, products, chart of accounts), (2) Financial Data (general ledger balances, open items, transaction history), (3) Human Resources Data (employee records, payroll history, benefits enrollment), (4) Supply Chain Data (inventory balances, purchase orders, sales orders), and (5) CRM Data (contacts, opportunities, cases, activities).",
    "For each phase, the migration process shall consist of the following steps: (a) Data Profiling - analyze source data to understand data quality, completeness, and transformation requirements; (b) Data Mapping - define the mapping rules between source and target data structures; (c) Transformation Rules - define the business rules for transforming, cleansing, and enriching source data; (d) ETL Development - develop and unit test the extraction, transformation, and loading routines; (e) Mock Migration - execute the migration in the staging environment with a representative data set; (f) Data Validation - validate the migrated data against source data and business rules; (g) Cutover Migration - execute the final migration in the production environment during the go-live window; and (h) Post-Migration Validation - perform comprehensive validation of all migrated data in the production environment.",
    "3. DATA QUALITY REQUIREMENTS",
    "All migrated data shall meet the following quality requirements: (a) Completeness - all required fields shall be populated with valid values; missing values shall be identified and resolved prior to migration; (b) Accuracy - migrated data shall match source data within the defined tolerance levels (zero percent (0%) tolerance for financial data, one percent (1%) tolerance for non-financial data); (c) Consistency - migrated data shall be consistent across all related tables and modules; referential integrity shall be maintained; (d) Timeliness - data shall be migrated within the defined cutover window; and (e) Uniqueness - duplicate records shall be identified and resolved prior to migration using the deduplication rules approved by Party A.",
    "4. DATA ARCHIVAL",
    "Historical data older than seven (7) years shall be archived rather than migrated to the new ERP system. Archived data shall be stored in a read-only data warehouse accessible through the Business Intelligence and Analytics module. The archive shall maintain full audit trail capabilities and shall comply with all applicable data retention requirements. Party B shall provide Party A with a detailed data archival report identifying all archived records, the source systems, and the archival dates.",
    "5. ROLLBACK PROCEDURES",
    "Party B shall develop and test detailed rollback procedures for each phase of the data migration. In the event that a migration phase fails or produces unacceptable results, Party B shall be capable of rolling back to the pre-migration state within four (4) hours. Rollback procedures shall be tested during each mock migration cycle. The rollback decision shall be made by the Steering Committee based on predefined criteria, including data quality metrics, system performance, and user feedback.",
  ]},
  { title: "SCHEDULE 5 - TRAINING PLAN", items: [
    "1. TRAINING APPROACH",
    "Party B shall develop and deliver a comprehensive training program designed to ensure that all Party A end users, administrators, and support personnel are proficient in using the new ERP system. The training program shall employ a blended learning approach that combines instructor-led training, hands-on workshops, e-learning modules, job aids, and ongoing performance support.",
    "2. TRAINING CURRICULUM",
    "The training curriculum shall be organized by user role and functional module, as follows:",
    "Executive Users: A four (4) hour training session covering dashboard navigation, report generation, key performance indicator monitoring, and strategic decision support capabilities. Target audience: C-suite executives, Vice Presidents, and Directors (approximately 50 users).",
    "Financial Users: A forty (40) hour training program covering all Financial Management module functionality, including general ledger operations, accounts payable processing, accounts receivable management, fixed asset accounting, cash management, month-end and year-end closing procedures, and financial reporting. Target audience: CFO, Controllers, Accountants, and Financial Analysts (approximately 75 users).",
    "HR Users: A thirty-two (32) hour training program covering all Human Capital Management module functionality, including employee lifecycle management, payroll processing, benefits administration, time and attendance, recruitment, onboarding, performance management, and compliance reporting. Target audience: CHRO, HR Business Partners, HR Specialists, and Payroll Administrators (approximately 60 users).",
    "Supply Chain Users: A thirty-six (36) hour training program covering all Supply Chain Management module functionality, including procurement, purchase order management, vendor management, inventory management, warehouse operations, order fulfillment, shipping and logistics, and supply chain analytics. Target audience: VP of Operations, Procurement Managers, Inventory Managers, Warehouse Supervisors, and Logistics Coordinators (approximately 100 users).",
    "Sales and Marketing Users: A twenty-four (24) hour training program covering all Customer Relationship Management module functionality, including contact and account management, opportunity tracking, pipeline management, campaign management, lead scoring, customer service case management, and sales analytics. Target audience: VP of Sales, Sales Managers, Account Executives, Marketing Managers, and Customer Service Representatives (approximately 150 users).",
    "System Administrators: A sixty (60) hour training program covering system configuration, user management, security administration, workflow configuration, integration management, system monitoring, performance tuning, backup and recovery procedures, and upgrade management. Target audience: IT Administrators and Support Staff (approximately 15 users).",
    "3. TRAINING DELIVERY SCHEDULE",
    "Training shall be delivered in three (3) waves: Wave 1 (Month 16) - Train-the-Trainer sessions for Party A's designated super users and training coordinators; Wave 2 (Month 17) - End user training for all user groups, delivered by a combination of Party B trainers and Party A super users; Wave 3 (Months 19-20) - Reinforcement training and advanced topics for users requiring additional support.",
    "4. TRAINING MATERIALS",
    "Party B shall develop the following training materials for each user role: (a) comprehensive user guides with step-by-step instructions and screenshots; (b) quick reference cards summarizing key procedures and navigation paths; (c) interactive e-learning modules with knowledge checks and assessments; (d) training videos demonstrating key system functions; (e) hands-on exercise workbooks with realistic scenarios and sample data; and (f) frequently asked questions documents addressing common issues and solutions.",
    "All training materials shall be delivered in both English and Spanish, shall comply with Section 508 accessibility requirements, and shall be maintained and updated by Party B for the duration of the Agreement. Party A shall have the right to reproduce and distribute training materials internally without restriction.",
    "5. TRAINING ENVIRONMENT",
    "Party B shall provision and maintain a dedicated training environment that mirrors the production environment in functionality and data. The training environment shall be refreshed with sanitized production data prior to each training wave. The training environment shall support at least one hundred (100) concurrent users to facilitate classroom-based training sessions. Party B shall ensure that the training environment is available during all scheduled training sessions and shall resolve any technical issues within one (1) hour.",
  ]},
];

for (const schedule of scheduleTopics) {
  bodyContent += wParagraph(schedule.title, { bold: true, size: 28, spacing: { before: 400, after: 300 }, keepNext: true });
  for (const item of schedule.items) {
    const isHeading = /^\d+\.\s+[A-Z]/.test(item);
    if (isHeading) {
      bodyContent += wParagraph(item, { bold: true, size: 24, spacing: { before: 300, after: 100 }, keepNext: true });
    } else {
      bodyContent += wParagraph(item, { size: 22, spacing: { before: 100, after: 100 } });
    }
  }
  bodyContent += pageBreak();
}

// Add appendix with additional legal terms to reach page target
const appendixSections = [];
for (let i = 1; i <= 30; i++) {
  const appendixParagraphs = [];
  for (let j = 0; j < 16; j++) {
    appendixParagraphs.push(legalPhrases[j % legalPhrases.length] + ` This provision shall apply specifically to the matters described in Appendix Section ${i}.${j + 1} and shall be interpreted in accordance with the general principles established in Article XI of the Agreement. The parties acknowledge that this provision has been negotiated at arm's length and reflects the mutual understanding and agreement of the parties with respect to the allocation of rights, obligations, and liabilities arising in connection with the subject matter of this Appendix.`);
  }
  appendixSections.push({ num: i, paragraphs: appendixParagraphs });
}

bodyContent += wParagraph("APPENDIX - SUPPLEMENTARY TERMS AND CONDITIONS", { bold: true, size: 28, spacing: { before: 400, after: 300 } });

for (const sec of appendixSections) {
  bodyContent += wParagraph(`Section A-${sec.num}. Supplementary Provisions`, { bold: true, size: 24, spacing: { before: 300, after: 100 }, keepNext: true });
  for (const para of sec.paragraphs) {
    bodyContent += wParagraph(para, { size: 22, spacing: { before: 100, after: 100 } });
  }
}

// Assemble full flat OPC XML
const documentXml = `<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="${R_NS}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:body>${bodyContent}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr></w:body></w:document>`;

const flatOpc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="${PKG_NS}">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="${REL_NS}">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      ${documentXml}
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

const outPath = path.join(__dirname, "..", "test-files", "Legal-Document-100-Pages.xml");
fs.writeFileSync(outPath, flatOpc, "utf-8");
console.log(`Written to: ${outPath}`);
console.log(`File size: ${(fs.statSync(outPath).size / 1024).toFixed(1)} KB`);
