const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, LevelFormat, PageBreak, BorderStyle
} = require('docx');
const fs = require('fs');

// ─── Helpers ───────────────────────────────────────────────────────────────

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial" })],
    spacing: { before: 360, after: 180 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F3864", space: 1 } }
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: "1F3864" })],
    spacing: { before: 280, after: 120 }
  });
}

function qHeading(num, marks, questionText) {
  return new Paragraph({
    children: [
      new TextRun({ text: `Q${num}. `, bold: true, size: 24, font: "Arial", color: "C00000" }),
      new TextRun({ text: questionText, bold: true, size: 24, font: "Arial" }),
      new TextRun({ text: `  [${marks} marks]`, bold: false, size: 22, font: "Arial", italics: true, color: "666666" })
    ],
    spacing: { before: 280, after: 100 },
    shading: { fill: "EEF2F7" }
  });
}

function ansLabel() {
  return new Paragraph({
    children: [new TextRun({ text: "Answer:", bold: true, size: 22, font: "Arial", color: "1F3864" })],
    spacing: { before: 80, after: 60 }
  });
}

function body(text, { bold = false, indent = false } = {}) {
  return new Paragraph({
    children: [new TextRun({ text, bold, size: 22, font: "Arial" })],
    spacing: { before: 60, after: 60 },
    alignment: AlignmentType.JUSTIFIED,
    indent: indent ? { left: 720 } : {}
  });
}

function subHeading(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 22, font: "Arial", underline: {} })],
    spacing: { before: 120, after: 60 }
  });
}

function bullet(text, ref = "bullets") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    children: [new TextRun({ text, size: 22, font: "Arial" })],
    spacing: { before: 40, after: 40 },
    alignment: AlignmentType.JUSTIFIED
  });
}

function spacer() {
  return new Paragraph({ children: [new TextRun("")], spacing: { before: 80, after: 80 } });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ─── Content ───────────────────────────────────────────────────────────────

const children = [];

// Title Page
children.push(
  new Paragraph({ children: [new TextRun({ text: "LAW352: TRADEMARKS AND ALLIED LAWS", bold: true, size: 40, font: "Arial", color: "1F3864" })], alignment: AlignmentType.CENTER, spacing: { before: 480, after: 240 } }),
  new Paragraph({ children: [new TextRun({ text: "Practice Question Bank — Complete Answer Key", size: 28, font: "Arial", color: "444444", italics: true })], alignment: AlignmentType.CENTER, spacing: { before: 0, after: 240 } }),
  new Paragraph({ children: [new TextRun({ text: "Session 2025–26", size: 24, font: "Arial", color: "666666" })], alignment: AlignmentType.CENTER, spacing: { before: 0, after: 720 } }),
  pageBreak()
);

// ─── PART A: 2-MARK QUESTIONS ─────────────────────────────────────────────
children.push(heading1("PART A — 2-Mark Questions (Short Answer)"), spacer());

// Q1
children.push(
  qHeading(1, 2, "What is a trademark? Name any two characteristics of a trademark."),
  ansLabel(),
  body("A trademark is a distinctive sign, symbol, word, device, logo, label, numerals, shape of goods, packaging, or combination of colours that is used by a trader or manufacturer to identify his goods or services and distinguish them from those of others. Under the Trade Marks Act, 1999, a trademark serves as a badge of origin. The mark allows consumers to associate the product with a particular source and to make informed purchasing decisions."),
  subHeading("Two key characteristics of a trademark:"),
  bullet("Distinctiveness: The mark must be capable of distinguishing the goods or services of one enterprise from those of another. It should not be generic, descriptive, or common in the trade."),
  bullet("Use in commerce: The mark must be used or intended to be used in relation to goods or services in the course of trade. A trademark derives its value from actual use in commercial activity."),
  spacer()
);

// Q2
children.push(
  qHeading(2, 2, "What is the primary objective of the Paris Convention in the context of trademark protection?"),
  ansLabel(),
  body("The Paris Convention for the Protection of Industrial Property, 1883, is a foundational international treaty administered by WIPO. Its primary objective in the context of trademark protection is to establish the principle of national treatment and the right of priority for trademark applicants across member countries."),
  body("The Convention ensures that nationals of any member state receive the same level of trademark protection in other member states as those countries grant to their own nationals. Critically, Article 4 of the Convention provides a priority right of six months, meaning that once an applicant files a trademark application in one member state, they can claim that same filing date ('priority date') in any other member state if they file within six months. This prevents competitors from registering the same mark in other countries during the window of international filing."),
  spacer()
);

// Q3
children.push(
  qHeading(3, 2, "Define 'distinctiveness' in trademark law. Why is it essential for trademark protection?"),
  ansLabel(),
  body("Distinctiveness in trademark law refers to the capacity of a mark to identify and distinguish the goods or services of one enterprise from those of others. A mark is said to be distinctive if consumers, upon seeing it, associate it with a particular source or origin rather than with the product category in general."),
  body("Distinctiveness is essential because it is the fundamental criterion for trademark protection under Section 9 of the Trade Marks Act, 1999. Without distinctiveness, a mark cannot perform its primary function — that of identifying the origin of goods. Generic and purely descriptive terms cannot be monopolised by one trader since they must remain available for all traders in that field. Distinctiveness prevents one party from unfairly monopolising common language while ensuring consumers are not misled about the origin of products."),
  spacer()
);

// Q4
children.push(
  qHeading(4, 2, "What is a generic mark? Give one example."),
  ansLabel(),
  body("A generic mark is a word or phrase that refers to the general category, class, or nature of the goods or services themselves, rather than indicating a specific brand or source. Generic terms cannot be registered as trademarks because they must remain freely available for all traders who deal in those goods or services. Registration of such marks would give one trader an unfair monopoly over common language."),
  body("Example: The word 'COMPUTER' for computers or 'BREAD' for bakery products are generic terms. No single trader can claim exclusive rights over such terms. Sometimes a previously valid trademark can become generic over time through improper use — a process called 'genericide'. Examples include 'Escalator' (formerly a Otis trademark) and 'Thermos'. In India, 'Xerox' is often used generically for photocopying, though Xerox Corporation actively fights genericide."),
  spacer()
);

// Q5
children.push(
  qHeading(5, 2, "What is 'deceptive similarity'? Give a brief example."),
  ansLabel(),
  body("'Deceptive similarity' refers to the degree of resemblance between two marks that is likely to cause confusion or deception in the minds of average consumers regarding the origin, source, or affiliation of the goods or services. Under Section 11 of the Trade Marks Act, 1999, a mark that is deceptively similar to an already registered mark is a ground for refusal of registration."),
  body("The test for deceptive similarity is that of the average consumer with imperfect recollection, who does not compare the two marks side-by-side but relies on memory."),
  body("Example: 'AMUL' and 'AMUL-YA' for dairy products would constitute deceptive similarity, as would 'COLGATE' and 'COLGATE-PLUS'. In the leading case Cadila Health Care Ltd. v. Cadila Pharmaceuticals Ltd. (2001), the Supreme Court held that in pharmaceutical marks, even a small degree of similarity can be deceptively similar given the public health risks."),
  spacer()
);

// Q6
children.push(
  qHeading(6, 2, "What rights does a registered trademark proprietor hold under the Trade Marks Act, 1999?"),
  ansLabel(),
  body("Under Section 28 of the Trade Marks Act, 1999, registration of a trademark confers the following rights on the proprietor:"),
  bullet("Exclusive right to use the registered trademark in relation to the goods or services for which it is registered."),
  bullet("Right to obtain relief in respect of infringement of the trademark in the manner provided by the Act."),
  bullet("Right to license the mark to others (registered user provisions under Section 48–54)."),
  bullet("Right to assign the mark with or without the goodwill of the business."),
  bullet("The mark can be used as a property asset for commercial and financial purposes."),
  body("Registration also creates a presumption of validity, and the proprietor can use the ® symbol. These rights are territorial and limited to the goods/services specified in the registration."),
  spacer()
);

// Q7
children.push(
  qHeading(7, 2, "Define trademark licensing. What is its basic purpose?"),
  ansLabel(),
  body("Trademark licensing is a contractual arrangement by which the owner of a trademark (the licensor) grants permission to another party (the licensee) to use the mark in relation to specified goods or services, subject to certain conditions, for a defined period and territory, usually in exchange for a royalty or fee."),
  body("The basic purpose of trademark licensing is to allow the trademark owner to commercially exploit the mark by expanding its reach without directly participating in all markets or business activities. It enables brand expansion through franchising, distribution networks, and co-branding while maintaining the proprietor's ownership of the mark. Under the Trade Marks Act, 1999, a licensee is recorded as a 'registered user' under Section 48 to ensure quality control and to maintain the distinctiveness of the mark."),
  spacer()
);

// Q8
children.push(
  qHeading(8, 2, "What is trademark infringement? State the key element required to prove it."),
  ansLabel(),
  body("Trademark infringement is the unauthorised use of a registered trademark (or a deceptively similar mark) in the course of trade in relation to goods or services for which the mark is registered, without the consent of the registered proprietor. It is governed by Sections 29 and 30 of the Trade Marks Act, 1999."),
  body("The key element required to prove infringement is: (a) the existence of a valid registered trademark, (b) unauthorised use of the mark or a mark that is identical or deceptively similar, (c) use in the course of trade in relation to the same or similar goods or services, and (d) likelihood of confusion or deception in the minds of consumers. The most critical element is the likelihood of confusion — the unauthorised mark must be likely to mislead consumers about the origin of the goods or services."),
  spacer()
);

// Q9
children.push(
  qHeading(9, 2, "What is the doctrine of exhaustion as a defense against trademark infringement?"),
  ansLabel(),
  body("The doctrine of exhaustion (also called the first-sale doctrine) holds that once a trademarked product is lawfully placed in commerce by the trademark owner or with their consent, the trademark owner's rights over that particular product are 'exhausted'. This means the owner cannot use trademark law to control the subsequent resale, redistribution, or use of that specific product."),
  body("As a defense to infringement under Section 30(3) of the Trade Marks Act, 1999, a person who purchases genuine trademarked goods and resells them cannot be held liable for infringement, provided the goods have not been materially altered. India follows the principle of international exhaustion — once goods are put on the market anywhere in the world by the proprietor or with consent, rights are exhausted. However, the defense fails if the defendant has altered the goods or if the resale is accompanied by conduct that damages the mark's reputation."),
  spacer()
);

// Q10
children.push(
  qHeading(10, 2, "What is passing off? State the three classic elements required to establish a passing off action."),
  ansLabel(),
  body("Passing off is a common law tort that protects the goodwill of an unregistered trademark owner from misrepresentation by another party. It allows a trader to prevent competitors from misrepresenting their goods or services as those of the claimant, thereby causing damage to the claimant's goodwill and reputation. In India, passing off is also recognized under Section 27(2) of the Trade Marks Act, 1999."),
  subHeading("The three classic elements (the 'classical trinity' from Reckitt & Colman v. Borden, 1990):"),
  bullet("Goodwill: The plaintiff must establish that their goods or business enjoy goodwill and reputation in the market, such that consumers associate the mark, get-up, or trade name with the plaintiff's business."),
  bullet("Misrepresentation: There must be a misrepresentation by the defendant — deliberate or innocent — that is likely to deceive consumers into believing that the defendant's goods or services are those of, or are connected with, the plaintiff."),
  bullet("Damage: The plaintiff must show actual or likely damage to their goodwill as a result of the defendant's misrepresentation."),
  spacer()
);

// Q11
children.push(
  qHeading(11, 2, "Distinguish between absolute grounds and relative grounds for refusal of trademark registration."),
  ansLabel(),
  body("Under the Trade Marks Act, 1999, grounds for refusal are classified into two categories:"),
  subHeading("Absolute Grounds (Section 9):"),
  body("These relate to the inherent nature of the mark itself. A mark is refused on absolute grounds if it: lacks distinctiveness; is exclusively descriptive of the character, quality, quantity, or purpose of the goods; consists of signs customary in the trade; or if it is contrary to public policy or morality, likely to deceive the public, contains scandalous matter, or is a shape resulting from the nature of the goods. These grounds apply regardless of whether a conflicting prior mark exists."),
  subHeading("Relative Grounds (Section 11):"),
  body("These arise from conflicts with pre-existing third-party rights. A mark is refused on relative grounds if it is identical or similar to an earlier registered trademark and the goods/services are identical or similar, causing a likelihood of confusion; or if it takes unfair advantage of or is detrimental to a well-known trademark. These are relative because they depend on whether a third party raises an objection during examination or opposition."),
  spacer()
);

// Q12
children.push(
  qHeading(12, 2, "What is the concept of 'acquired distinctiveness'? How does it affect registration of an otherwise non-registrable mark?"),
  ansLabel(),
  body("Acquired distinctiveness (also called 'secondary meaning' or 'distinctiveness through use') refers to the process by which a mark that is prima facie non-distinctive, descriptive, or even generic acquires the ability to identify the source of goods through long and extensive use in the market. Over time, consumers associate the mark exclusively with a particular trader's goods, notwithstanding its original descriptive character."),
  body("Under the proviso to Section 9(1) of the Trade Marks Act, 1999, a trademark that would otherwise be refused registration on the ground that it lacks distinctiveness or is descriptive may still be registered if the applicant demonstrates that the mark has acquired distinctiveness before the date of application through actual use. Evidence such as sales figures, advertising expenditure, consumer surveys, duration of use, and market recognition is considered. Thus, 'acquired distinctiveness' serves as a statutory exception that allows otherwise non-registrable marks to gain protection."),
  spacer()
);

// Q13
children.push(
  qHeading(13, 2, "What is the difference between assignment and licensing of a trademark?"),
  ansLabel(),
  body("Assignment and licensing are both mechanisms to transfer rights in a trademark, but they differ fundamentally:"),
  subHeading("Assignment (Sections 37–45):"),
  body("Assignment involves a permanent transfer of ownership of the trademark from the assignor to the assignee. The assignor divests all or part of their rights in the mark. Assignment may be with goodwill (the assignee also acquires the brand value associated with the mark) or without goodwill (the assignee gets the mark but not the business reputation). After assignment, the assignor has no further right to use the mark."),
  subHeading("Licensing (Sections 48–54):"),
  body("Licensing involves only a grant of permission to use the trademark, while ownership remains with the licensor. The relationship is contractual and temporary. The licensee (registered user) uses the mark under the licensor's supervision and quality control obligations. The licensor retains ownership and the mark must not be used in a manner that deceives the public. Unlike assignment, licensing does not transfer goodwill or ownership."),
  spacer()
);

// Q14
children.push(
  qHeading(14, 2, "Distinguish between 'descriptive use' and 'unauthorized use' as defenses to trademark infringement."),
  ansLabel(),
  body("Descriptive use (fair use defense) under Section 30(2)(a) of the Trade Marks Act, 1999, permits a third party to use a registered trademark in a descriptive, non-trademark manner. This applies when the use is of the person's own name, address, or a description of the character or quality of the goods/services, provided such use is honest and in accordance with accepted practices. It is a bona fide use that does not leverage the mark as an indicator of origin. Example: A company using 'GOLD' to describe the quality of its product, where 'GOLD' is someone's trademark, may claim descriptive use if it is used in a purely descriptive sense."),
  body("Unauthorized use on the other hand refers to use of the registered mark without the owner's consent in a trademark sense — i.e., to indicate origin of goods. This constitutes infringement under Section 29. The distinction lies in the nature of use: descriptive use is non-source-identifying honest use, while unauthorized use appropriates the mark's brand function and is commercially exploitative."),
  spacer()
);

// Q15
children.push(
  qHeading(15, 2, "What is a geographical indication? How is it different from a trademark?"),
  ansLabel(),
  body("A Geographical Indication (GI) is a sign used on products that have a specific geographical origin and possess qualities, reputation, or characteristics essentially attributable to that place of origin. In India, GIs are protected under the Geographical Indications of Goods (Registration and Protection) Act, 1999. Examples include Darjeeling Tea, Kancheepuram Silk, Alphonso Mangoes, and Basmati Rice."),
  subHeading("Key differences from a trademark:"),
  bullet("Ownership: A trademark is owned by an individual or company. A GI is a collective right — it belongs to all producers in the designated region who meet the specified standards."),
  bullet("Transferability: Trademarks can be assigned or licensed to any party. A GI cannot be transferred, assigned, or licensed to parties outside the specified region."),
  bullet("Purpose: A trademark identifies the commercial source of goods. A GI identifies geographical origin and qualities linked to that origin."),
  bullet("Duration: GIs can be renewed indefinitely as long as the association with the region persists, similar to trademarks."),
  spacer()
);

// Q16
children.push(
  qHeading(16, 2, "How are domain name disputes resolved under the UDRP framework?"),
  ansLabel(),
  body("The Uniform Domain Name Dispute Resolution Policy (UDRP) is administered by ICANN and provides a streamlined arbitration mechanism for resolving disputes between trademark owners and domain name registrants. It applies to all generic top-level domains (.com, .net, .org, etc.)."),
  body("Under UDRP, a complainant must prove three elements: (1) the domain name is identical or confusingly similar to a trademark or service mark in which the complainant has rights; (2) the registrant has no legitimate interests in the domain name; and (3) the domain was registered and is being used in bad faith. Bad faith indicators include: registration to sell the domain to the trademark owner at a profit, intent to disrupt a competitor's business, and deliberate attempt to attract users for commercial gain by creating confusion. Proceedings are conducted online before approved dispute resolution service providers (WIPO, NAF, etc.) and are typically concluded within 45–60 days. Remedies are limited to cancellation or transfer of the domain; no monetary damages are awarded."),
  spacer()
);

// Q17
children.push(
  qHeading(17, 2, "In what circumstances can the use of a registered trademark itself be a defense to an infringement action? Explain with reference to Section 30."),
  ansLabel(),
  body("Section 30 of the Trade Marks Act, 1999, carves out situations where even what appears to be use of a registered trademark does not constitute infringement. These statutory defenses are available to any person:"),
  bullet("Section 30(2)(a): Use of a mark in a bona fide descriptive manner — to indicate the kind, quality, quantity, intended purpose, value, geographical origin, or time of production of goods or services — is not infringement."),
  bullet("Section 30(2)(b): Use of a trademark to indicate the intended purpose of a product or service (e.g., spare parts or accessories compatible with another's branded product) is permissible if the use is necessary and not misleading."),
  bullet("Section 30(3): Where a registered trademark has been applied to goods lawfully put on the Indian market, the proprietor's rights in respect of those goods are exhausted, and subsequent dealings in those goods do not constitute infringement."),
  body("Additionally, the use of one's own registered trademark is a defense against infringement claims by the proprietor of another registered mark. The underlying principle of Section 30 is that trademark rights must not be used to stifle legitimate descriptive, comparative, or resale use of goods or marks in commerce."),
  spacer()
);

// Q18
children.push(
  qHeading(18, 2, "Can publicity rights be enforced against trademark registrations? Briefly analyze."),
  ansLabel(),
  body("Publicity rights (the right of an individual — typically a celebrity — to control the commercial use of their name, image, likeness, or persona) exist at the intersection of personality rights, privacy law, and intellectual property law. India does not have a dedicated statute for publicity rights; however, courts have recognized them through judicial pronouncements drawing on the right to privacy under Article 21 of the Constitution, passing off principles, and tort law."),
  body("In ICC Development v. Arvee Enterprises (2003), the Delhi High Court recognized that the right of publicity vests in an individual and is actionable against unauthorized commercial exploitation. In Titan Industries v. Ramkumar Jewellers (2012), the court held that the image rights of celebrities could be enforced. If someone registers a trademark that incorporates a celebrity's name, face, or signature without consent, the celebrity can challenge the registration on the ground that it takes unfair advantage under Section 11(3) (well-known mark protection) or invoke passing off. However, enforcement is challenging due to absence of a standalone statute, and remedies are typically injunctions and damages through civil courts."),
  spacer()
);

// Q19
children.push(
  qHeading(19, 2, "What are the limitations on the scope of trademark licensing that prevent the mark from becoming deceptive?"),
  ansLabel(),
  body("The primary concern with trademark licensing is that if the licensor does not control the quality of goods produced by the licensee, the mark may mislead consumers — since consumers rely on the mark as an assurance of consistent quality from a single source. To prevent this, the Trade Marks Act, 1999, imposes the following limitations:"),
  bullet("Quality control obligation: The licensor must exercise adequate supervision and control over the use of the mark by the licensee, particularly with respect to the quality of goods or services. This is implicit in the registered user provisions."),
  bullet("Prohibition on naked licensing: A license without any quality control (called a 'naked license') is impermissible as it renders the mark deceptive. The mark may face cancellation if the public is misled."),
  bullet("Record of registered user: Under Section 48, a licensee must be registered as a 'registered user' to have legal standing, ensuring the Registrar is aware of the licensing arrangement."),
  bullet("Territorial and purpose limitations: A license can be restricted to a specific territory, duration, or type of goods, preventing the licensee from using the mark beyond the agreed scope."),
  body("These limitations ensure the mark continues to serve its fundamental function as a reliable indicator of origin and quality."),
  spacer()
);

// Q20
children.push(
  qHeading(20, 2, "How does the Madrid System differ from the Paris Convention route for international trademark registration? Briefly compare."),
  ansLabel(),
  body("Both the Madrid System and the Paris Convention facilitate international trademark protection, but they operate through different mechanisms:"),
  subHeading("Paris Convention Route:"),
  body("The Paris Convention (1883) provides a right of priority of six months — a trademark applicant filing in one member country can claim the same filing date in other member countries if they file within six months. However, the applicant must file separate national applications in each country, comply with each country's procedural requirements, and pay separate fees to each national office. There is no centralized registration; the Convention only harmonises standards and grants priority rights."),
  subHeading("Madrid System (Madrid Agreement + Madrid Protocol):"),
  body("The Madrid System, administered by WIPO, allows an applicant to file a single international application in one language (English, French, or Spanish) and pay a single fee set of fees to obtain trademark protection in up to 130+ countries. The application is based on a home-country 'basic mark'. Each designated country examines the application under its own law, but the process is administratively centralized. India acceded to the Madrid Protocol in 2013."),
  subHeading("Key difference:"),
  body("The Paris Convention is a framework of principles and priority rights; the Madrid System is an operational filing system that actually enables centralized, single-application, multi-country registration."),
  spacer()
);

// ─── PART B: 10-MARK QUESTIONS ────────────────────────────────────────────
children.push(pageBreak(), heading1("PART B — 10-Mark Questions (Long Answer)"), spacer());

// B-Q1
children.push(
  qHeading(1, 10, "Trace the evolution of trademark law in India from colonial-era legislation to the Trade Marks Act, 1999. Discuss the key changes brought about by the 1999 Act in terms of scope and coverage of trademarks."),
  ansLabel(),
  subHeading("1. Introduction"),
  body("Trademark law in India has undergone significant transformation over more than a century, evolving from rudimentary colonial-era statutory provisions to a comprehensive modern legislative framework aligned with international standards. The journey from the Trade Marks Act, 1940 through the Trade and Merchandise Marks Act, 1958 to the Trade Marks Act, 1999 reflects India's growing commitment to intellectual property protection and its integration into the global trading order."),
  subHeading("2. Colonial Era: Pre-Independence Development"),
  body("Prior to the enactment of formal trademark legislation, trademark protection in India was provided under the common law doctrine of passing off, which was inherited from English jurisprudence. Traders seeking to protect their marks against imitation had to resort to civil courts on principles of equity and tort. The Transfer of Property Act, 1882, and the Contract Act, 1872, provided incidental protection to trade names and marks in commercial transactions but fell short of a dedicated trademark regime."),
  body("The first dedicated trademark statute in India was the Trade Marks Act, 1940, modelled on the English Trade Marks Act, 1938. This was a pioneering piece of legislation that established a formal system of registration, created the Trade Marks Registry, and provided for legal remedies for infringement. However, its scope was limited — it covered only goods (not services) and did not deal with collective marks, certification marks, or well-known marks in any systematic way."),
  subHeading("3. The Trade and Merchandise Marks Act, 1958"),
  body("After independence, India enacted the Trade and Merchandise Marks Act, 1958, which replaced the 1940 Act and brought greater clarity and procedural detail. The 1958 Act consolidated the law on trademarks and merchandise marks and provided a more structured framework for registration, assignment, licensing, and infringement. However, it still had significant limitations:"),
  bullet("It did not cover service marks — marks used in connection with services rather than goods."),
  bullet("It did not address the increasingly important issue of well-known marks and their cross-category protection."),
  bullet("International harmonisation was limited; India was not a party to the Madrid Protocol at the time."),
  bullet("The Act did not define or protect collective marks and certification marks adequately."),
  bullet("It did not account for the digital economy and internet-related trademark issues."),
  subHeading("4. India's International Obligations and the Need for Reform"),
  body("The most significant catalyst for modernizing Indian trademark law was India's becoming a signatory to the Agreement on Trade-Related Aspects of Intellectual Property Rights (TRIPS), 1994, as part of its WTO accession. TRIPS mandated comprehensive minimum standards of intellectual property protection, including for trademarks. India was required to extend trademark protection to service marks, provide enhanced protection for well-known trademarks, and align its procedural and enforcement framework with international standards."),
  subHeading("5. The Trade Marks Act, 1999 — Key Changes"),
  body("The Trade Marks Act, 1999, which came into force on September 15, 2003, represents a watershed in Indian trademark jurisprudence. The following are its major contributions:"),
  body("(a) Extension to Service Marks: For the first time, the 1999 Act expressly extended trademark protection to service marks, recognizing marks used in the provision of services such as banking, insurance, telecommunication, and hospitality. This brought India in line with TRIPS requirements."),
  body("(b) Well-Known Trademarks: The 1999 Act introduced enhanced protection for 'well-known trademarks' (Section 2(1)(zg) and Section 11). A well-known mark is entitled to protection even for goods and services that are dissimilar to those for which the mark is registered, addressing the issue of dilution."),
  body("(c) Collective Marks and Certification Marks: The Act provided for the registration of collective marks (used by associations of producers, traders, or service providers) and certification marks (attesting to quality or origin of goods), expanding the types of protectable marks."),
  body("(d) 3D Marks and Non-Traditional Marks: The 1999 Act broadened the definition of a trademark to include the shape of goods, their packaging, and combination of colours, thereby recognizing three-dimensional and non-traditional trademarks."),
  body("(e) Opposition and Cancellation Procedures: More detailed provisions for opposition to registration and rectification/cancellation of registered marks were introduced, strengthening the system of quality control in the Register."),
  body("(f) Enhanced Penalties and Criminal Remedies: Criminal sanctions were significantly enhanced with increased imprisonment and fines for trademark counterfeiting and infringement, reflecting the seriousness with which the law treats trademark fraud."),
  body("(g) Madrid Protocol Alignment: While India acceded to the Madrid Protocol only in 2013, the 1999 Act laid the groundwork by harmonizing the registration system with international norms."),
  subHeading("6. Recent Amendments and Developments"),
  body("The Trade Marks (Amendment) Act, 2010 made further changes to facilitate India's accession to the Madrid Protocol. The Trade Marks Rules, 2017 overhauled procedural aspects, introduced e-filing, and reduced government fees for startups. Indian courts have also progressively developed the law through landmark decisions such as Amritdhara Pharmacy v. Satya Deo (1963), Cadila Health Care v. Cadila Pharmaceuticals (2001), and Toyota Jidosha v. Prius Auto Industries (2017)."),
  subHeading("7. Conclusion"),
  body("The evolution from the 1940 Act through the 1958 Act to the 1999 Act illustrates India's progressive response to domestic commercial realities, international trade obligations, and technological change. The Trade Marks Act, 1999 significantly expanded the scope and coverage of trademarks, aligned Indian law with TRIPS and WIPO standards, and equipped courts and the Registry with modern tools to protect trademark rights in an increasingly complex commercial environment."),
  spacer()
);

// B-Q2
children.push(
  qHeading(2, 10, "Explain the distinctiveness spectrum in trademark law. Using relevant case law, discuss how marks ranging from generic to arbitrary/fanciful are treated under Indian trademark law."),
  ansLabel(),
  subHeading("1. Introduction"),
  body("The concept of distinctiveness lies at the very heart of trademark law. Only marks that are capable of distinguishing the goods or services of one trader from those of others merit legal protection. The distinctiveness spectrum — ranging from generic at one end to arbitrary/fanciful at the other — provides a structured framework for evaluating the protectability of trademarks. The more distinctive a mark, the stronger its legal protection; the more generic or descriptive, the weaker or non-existent its protection."),
  subHeading("2. The Spectrum Explained"),
  body("The distinctiveness spectrum, derived from the influential US case Abercrombie & Fitch Co. v. Hunting World Inc. (1976) and incorporated into Indian trademark jurisprudence, classifies marks into five categories:"),
  subHeading("(a) Generic Marks"),
  body("Generic marks are the common name or genus of the product or service itself. They are completely devoid of distinctiveness because they describe what the product is rather than where it comes from. No single trader can monopolise common language. Under Section 9(1)(a) of the Trade Marks Act, 1999, such marks cannot be registered."),
  body("Example: 'COMPUTER' for computers, 'SUGAR' for sugar, 'SOAP' for soap. Indian courts have consistently held that generic terms cannot function as trademarks. Generic marks can arise either naturally (e.g., 'BREAD') or through 'genericide', where a formerly valid trademark degenerates into a generic term through widespread misuse by the public."),
  subHeading("(b) Descriptive Marks"),
  body("Descriptive marks are those that directly describe the quality, character, function, geographical origin, or other features of the goods or services. They are not inherently distinctive and cannot be registered under Section 9(1)(b) of the Act unless they have acquired 'secondary meaning' through long and extensive use (Section 9(1) proviso)."),
  body("Example: 'SWEET' for candies, 'COLD AND CREAMY' for ice cream. In Harrods Ltd. v. Harrodian School Ltd. (UK, applied in India), descriptiveness was considered a bar to registration. Indian courts in cases like Pankaj Goel v. Dabur India (2008) have examined the descriptive character of marks in pharmaceutical contexts."),
  subHeading("(c) Suggestive Marks"),
  body("Suggestive marks indirectly suggest some quality or characteristic of the goods but require consumer imagination to connect the mark to the product. They are inherently distinctive and registrable without proof of use. They occupy the middle ground on the spectrum."),
  body("Example: 'PENGUIN' for books (suggesting a friendly, literary character), 'NETFLIX' (suggesting 'net' and a flux of content), 'CITIBANK' (suggesting city and banking). In India, 'GLUCON-D' was treated as suggestive of glucose energy in Heinz Italia v. Dabur India (2007), balancing distinctiveness with descriptive overtones."),
  subHeading("(d) Arbitrary Marks"),
  body("Arbitrary marks are real words with a well-known meaning that are applied to products with which they have no logical connection whatsoever. They are inherently distinctive and entitled to strong protection."),
  body("Example: 'APPLE' for computers, 'CAMEL' for cigarettes (no connection between camels and tobacco), 'LOTUS' for software. Indian courts readily grant injunctions in favor of such marks. In Apple Inc. v. Apple Computers (India), the mark's arbitrary nature was a key factor in establishing strong rights."),
  subHeading("(e) Fanciful/Invented Marks"),
  body("Fanciful or invented marks are coined words that have no prior meaning in any language. They are the most inherently distinctive category and receive the strongest legal protection. They exist solely as trademarks and have no meaning other than to identify a brand."),
  body("Example: 'KODAK' (coined by George Eastman), 'XEROX', 'PEPSI', 'ZOMATO'. In India, 'FEVICOL' (Pidilite), 'BORO PLUS', and 'AMUL' (acronym, but functioning as a fanciful mark) enjoy strong protection. In Amul v. Amul-Ya, the court recognized the strength of the invented mark in enjoining an infringing use."),
  subHeading("3. Acquired Distinctiveness / Secondary Meaning"),
  body("As noted, descriptive marks can 'move up' the spectrum by acquiring distinctiveness through use. Indian courts consider factors such as duration and geographical extent of use, advertising expenditure, consumer surveys, and media coverage. In Ranjit Udyog v. Karan Udyog (2005), the court held that even a descriptive mark can acquire distinctiveness over time through extensive market exposure."),
  subHeading("4. Judicial Approach in India"),
  body("Indian courts apply the 'unwary purchaser' test and the 'imperfect recollection' test when evaluating trademark distinctiveness. In Cadila Health Care v. Cadila Pharmaceuticals (2001), the Supreme Court laid down comprehensive guidelines for evaluating similarity and distinctiveness in pharmaceutical marks, noting that even phonetically or visually similar marks may confuse consumers. In Parle Products v. J.P. & Co. (1972), the Supreme Court emphasised that the overall impression of the mark, not microscopic analysis, determines distinctiveness."),
  subHeading("5. Conclusion"),
  body("The distinctiveness spectrum is an essential analytical tool in trademark law. Indian courts and the Trade Marks Registry consistently apply this framework to determine the registrability and protectability of marks. The stronger and more distinctive a mark — especially at the arbitrary and fanciful end of the spectrum — the greater the protection it commands. The law thereby incentivises creative branding while preventing monopolisation of common language."),
  spacer()
);

// B-Q3
children.push(
  qHeading(3, 10, "Describe the step-by-step procedure for registration of a trademark in India. What pre-requisites must be fulfilled before filing an application?"),
  ansLabel(),
  subHeading("1. Introduction"),
  body("Registration of a trademark under the Trade Marks Act, 1999 (as amended by the Trade Marks (Amendment) Act, 2010, and the Trade Marks Rules, 2017) provides the proprietor with exclusive rights and a statutory presumption of ownership. The registration process involves several well-defined stages from the initial search to the final entry on the Register."),
  subHeading("2. Pre-Requisites Before Filing"),
  body("Before filing a trademark application, the following conditions must be assessed:"),
  bullet("The mark must be a 'trademark' as defined in Section 2(1)(zb) — capable of graphical representation and capable of distinguishing goods or services of one person from those of others."),
  bullet("The mark must not fall under the absolute grounds of refusal under Section 9 (lack of distinctiveness, descriptive, generic, deceptive, against public policy, etc.)."),
  bullet("The applicant must be the proprietor of the mark — either through prior use or through first filing (in case of no prior use)."),
  bullet("The goods or services must be identified and classified under the appropriate Nice Classification (1–45 classes), as India follows the Nice Classification System."),
  bullet("A preliminary trademark search should be conducted on the Trademark Registry's online database (IP India website) to ensure no identical or deceptively similar mark is already registered or pending for the same/similar goods and services."),
  subHeading("3. Step-by-Step Registration Procedure"),
  subHeading("Step 1: Trademark Search (Pre-Filing)"),
  body("A comprehensive search of the Trade Marks Register should be conducted to check for conflicting marks. While not mandatory, it is strongly advisable to avoid opposition and rejection. The search can be done online through the IP India Trade Marks Registry portal (tmrsearch.ipindia.gov.in) or through an attorney."),
  subHeading("Step 2: Filing the Application"),
  body("The application is filed with the Trade Marks Registry under Section 18 of the Act. The application must include: the applicant's details; a clear representation of the trademark (logo, word, device, etc.); the specification of goods/services with the applicable class(es) under the Nice Classification; the date of first use (if the mark has been used prior to filing) or a statement of proposed use; the prescribed fee (as per the Trade Marks Rules, 2017 — reduced for startups and individuals). Applications can be filed online (Form TM-A) through the IP India e-filing portal or physically at one of the five trademark offices (Mumbai, Delhi, Chennai, Kolkata, Ahmedabad)."),
  subHeading("Step 3: Examination"),
  body("After filing, the Registry assigns the application a serial number and allots it to an Examiner. The Examiner reviews the application for: formalities and compliance; absolute grounds of refusal under Section 9; relative grounds of refusal under Section 11 (conflicts with existing registrations or pending applications). The examination report is typically issued within 30–90 days of filing. If objections are raised, the applicant receives a show-cause notice."),
  subHeading("Step 4: Response to Examination Report / Hearing"),
  body("If the Examiner raises objections, the applicant must file a response within one month of receiving the examination report (extendable by one further month on application). If the written response does not resolve the objections, a personal hearing before the Registrar or Examiner is scheduled. The applicant argues why the mark should be accepted despite the objections."),
  subHeading("Step 5: Acceptance and Advertisement"),
  body("If the Registrar is satisfied with the application (either on original filing or after response and hearing), the trademark is accepted for advertisement. The mark is published in the Trade Marks Journal (published weekly on the IP India website), inviting public opposition. This is a critical stage as it opens the mark to challenges from third parties."),
  subHeading("Step 6: Opposition Proceedings"),
  body("Any person may oppose the registration within four months of the date of advertisement in the Trade Marks Journal (Section 21). The Opponent files a Notice of Opposition (Form TM-O) with the prescribed fee. The applicant files a Counter Statement (Form TM-O) within two months. Evidence is then filed by both sides in stages (Opposition Board may be involved). Hearings are conducted, and the Registrar decides whether the mark should proceed to registration or be refused."),
  subHeading("Step 7: Registration and Issue of Certificate"),
  body("If no opposition is filed within the four-month window, or if opposition is decided in favor of the applicant, the Registrar enters the trademark in the Register of Trade Marks and issues a Certificate of Registration (Form TM-23). The registration is effective from the date of the original application (not the date of registration certificate). The ® symbol may be used only after registration."),
  subHeading("Step 8: Renewal"),
  body("A registered trademark is valid for ten years from the date of application and is indefinitely renewable for successive periods of ten years each on payment of the prescribed renewal fee under Section 25. A trademark must be renewed before expiry; a six-month grace period is available for renewal with a surcharge."),
  subHeading("4. Conclusion"),
  body("The registration process under the Trade Marks Act, 1999 is designed to balance the interests of applicants, existing trademark owners, and the public. The multi-stage process ensures that only marks that genuinely merit protection receive it, while providing third parties ample opportunity to raise legitimate objections. Online filings and the Trademark e-portal have significantly streamlined the process in recent years."),
  spacer()
);

// B-Q4
children.push(
  qHeading(4, 10, "Explain the concept of trademark licensing under the Trade Marks Act, 1999. What are the legal requirements for a valid license and what limitations does the law impose on the licensor and licensee?"),
  ansLabel(),
  subHeading("1. Introduction"),
  body("Trademark licensing is one of the most commercially significant mechanisms in modern intellectual property law, enabling trademark owners to monetise their marks and expand their commercial reach without directly participating in every market. The Trade Marks Act, 1999 governs trademark licensing primarily through Sections 48 to 54, which deal with 'registered users', and through the general framework of the Act which recognizes licensing as a lawful disposition of trademark rights. Licensing is the cornerstone of franchising, distribution networks, merchandising agreements, and global brand expansion strategies."),
  subHeading("2. Concept of Trademark Licensing"),
  body("A trademark license is a contractual arrangement whereby the proprietor of a registered trademark (the licensor) grants permission to another party (the licensee or 'registered user') to use the mark in relation to specified goods or services, in a defined territory and for a defined period, typically in exchange for royalty payments. The licensor retains ownership of the trademark, and the license can be exclusive (only the licensee and no other, possibly not even the licensor), sole (licensor and licensee only), or non-exclusive (licensor can grant to multiple licensees)."),
  body("In India, a licensee who wishes to have legal standing to sue for infringement must be recorded as a 'registered user' under Section 48 of the Act. An unregistered licensee does not lose the right to use the mark, but may have limited standing in infringement proceedings."),
  subHeading("3. Legal Requirements for a Valid License"),
  subHeading("(a) Written Agreement"),
  body("While the Act does not categorically require a written license agreement (unlike assignment, which must be in writing under Section 37), it is strongly advisable to have a written license agreement to define terms, prevent disputes, and serve as evidence before courts and the Registry."),
  subHeading("(b) Application for Registered User Status (Section 48)"),
  body("The proprietor and the proposed licensee must jointly apply to the Registrar (using Form TM-U) to have the licensee recorded as a 'registered user'. The application must contain: particulars of the relationship between proprietor and registered user; the goods or services for which the mark will be used; the conditions and restrictions (if any); and a declaration of consent from the proprietor."),
  subHeading("(c) Specification of Terms"),
  body("The license must clearly specify: the territory within which the mark may be used; the duration of the license; the goods or services covered; any quality standards the licensee must comply with; the royalty or other consideration (if any); and provisions for termination and renewal."),
  subHeading("(d) Quality Control"),
  body("A fundamental requirement — though not always explicitly stated in the Act — is that the licensor must maintain adequate quality control over the goods or services produced by the licensee under the mark. This is critical because the trademark serves as a guarantee of quality to consumers. A failure to maintain quality control may result in the mark becoming deceptive or generic, and could be grounds for cancellation."),
  subHeading("4. Limitations Imposed on Licensor and Licensee"),
  subHeading("(a) No Naked Licensing"),
  body("The licensor cannot issue a 'naked license' — i.e., a license without exercising any quality control over the licensee's goods or services. If the proprietor allows use of the mark without any oversight, the mark may mislead consumers into associating the mark with a consistent quality that does not exist, rendering the mark deceptive. Courts have the power to cancel such registrations under Section 57."),
  subHeading("(b) Prohibition on Deceptive Use"),
  body("Under Section 50, the Registrar can cancel a registered user's rights if the use of the mark by the licensee is likely to cause deception or confusion in the public. The licensee must use the mark only in accordance with the license terms and must not use it in a manner that degrades the mark or misleads consumers."),
  subHeading("(c) Territorial Restrictions"),
  body("A license can be territorially limited, and the licensee must not use the mark outside the licensed territory. This is particularly important in sub-licensing scenarios, where the original licensee seeks to further sub-license the mark, which requires express permission from the proprietor."),
  subHeading("(d) Parallel Use and Confusion"),
  body("Where multiple registered users exist for the same mark in different territories, their use of the mark must not result in public confusion. The Registrar can refuse registration of a proposed registered user if the arrangement would be contrary to public interest."),
  subHeading("(e) Quality Control Obligations"),
  body("Section 49 requires the application for registered user to contain details of the conditions and restrictions, and the nature of the control the proprietor will exercise. The Act implicitly requires that the proprietor supervise the licensee's use of the mark to ensure the quality and nature of the goods are consistent with the mark's identity."),
  subHeading("5. Judicial Treatment"),
  body("Indian courts have consistently upheld the importance of quality control in licensing. In cases involving franchise arrangements — like those in the food and beverage sector (McDonald's, KFC, Domino's) — courts have recognized the licensee's use as the proprietor's use, provided quality standards are maintained. The Supreme Court in several passing off matters has noted that trademark licensing must not be a vehicle for consumer deception."),
  subHeading("6. Conclusion"),
  body("Trademark licensing under the Trade Marks Act, 1999 is a carefully regulated mechanism that balances commercial freedom with consumer protection. The law insists on quality control, transparency through registered user procedures, and limitations on territorial and functional use to ensure the mark continues to fulfil its primary function — indicating origin and assuring quality. With the growth of global franchising and digital commerce, the law of trademark licensing is increasingly relevant and important."),
  spacer()
);

// B-Q5
children.push(
  qHeading(5, 10, "What constitutes infringement of a registered trademark under the Trade Marks Act, 1999? Discuss the role of 'likelihood of confusion' in determining infringement, with reference to relevant case law."),
  ansLabel(),
  subHeading("1. Introduction"),
  body("Trademark infringement is the unauthorized use of a registered trademark in a manner that violates the exclusive rights of the registered proprietor. The Trade Marks Act, 1999 contains a detailed statutory framework governing infringement in Sections 29 and 30. Central to any infringement analysis is the question of whether the impugned use creates a 'likelihood of confusion' in the minds of consumers — a standard that Indian courts have elaborated upon through decades of jurisprudence."),
  subHeading("2. What Constitutes Infringement — Section 29"),
  body("Section 29 of the Trade Marks Act, 1999 provides an exhaustive list of acts that constitute trademark infringement:"),
  bullet("Section 29(1): A person infringes a registered trademark if they use in the course of trade a mark that is identical with or deceptively similar to the registered mark in relation to goods or services for which the mark is registered, and such use is likely to cause confusion in the minds of consumers (including a likelihood of association with the registered mark)."),
  bullet("Section 29(2): Where an identical/similar mark is used for identical/similar goods or services, and where there is a likelihood of confusion, infringement is established. If the marks are identical and goods are identical, confusion is presumed."),
  bullet("Section 29(4): Infringement also occurs where the infringing mark is identical or similar to a well-known registered mark and is used for dissimilar goods or services, if such use takes unfair advantage of or is detrimental to the distinctive character or repute of the registered mark (dilution)."),
  bullet("Section 29(6): The following uses also constitute infringement: affixing the mark on goods/packaging; offering goods for sale or supply under the mark; importing/exporting goods with the mark; using the mark on business papers or in advertising."),
  subHeading("3. Essential Elements of Infringement"),
  body("To establish infringement under Section 29, the plaintiff must prove: (i) ownership of a valid registered trademark; (ii) that the defendant has used a mark identical or deceptively similar to the plaintiff's mark; (iii) that the use is in the course of trade; (iv) that the use is in relation to identical or similar goods or services (except in cases of well-known marks where this requirement is relaxed); and (v) that there is a likelihood of confusion or deception among consumers."),
  subHeading("4. The Role of 'Likelihood of Confusion'"),
  body("The concept of 'likelihood of confusion' is the cornerstone of trademark infringement analysis. It captures the probability that consumers will be misled about the source, origin, affiliation, or sponsorship of goods or services. Courts do not require evidence of actual confusion — a reasonable likelihood of confusion is sufficient."),
  subHeading("Factors Considered by Indian Courts:"),
  bullet("Strength of the mark: A highly distinctive or well-known mark enjoys broader protection; a descriptive or weak mark gets narrower protection."),
  bullet("Similarity of marks: Visual, phonetic, and conceptual similarity are all considered. Courts apply the 'anti-dissection rule' — marks must be compared in their entirety, not element by element."),
  bullet("Similarity of goods/services: The closer the goods or services in the marketplace, the greater the risk of confusion. Confusion is presumed when both marks and goods are identical (Section 29(3))."),
  bullet("Consumer profile: The standard is the average consumer with imperfect recollection, not an expert comparing marks side by side."),
  bullet("Evidence of actual confusion: While not required, evidence of actual consumer confusion is highly persuasive."),
  subHeading("5. Leading Case Law"),
  body("Cadila Health Care Ltd. v. Cadila Pharmaceuticals Ltd. (2001 SC): The Supreme Court laid down comprehensive guidelines for determining infringement in pharmaceutical marks, emphasizing that even phonetically similar marks for medicinal products can cause grave confusion given the vulnerability of consumers. The Court listed factors including nature of the mark, the degree of resemblance, nature of goods, class of consumers, and evidence of actual confusion."),
  body("Amritdhara Pharmacy v. Satya Deo Gupta (1963 SC): The Supreme Court established the 'average man of imperfect recollection' test, holding that marks must be compared as a whole and not by dissecting them. 'AMRITDHARA' and 'LAKSHMANDHARA' were held to be confusingly similar for medicinal preparations."),
  body("Parle Products v. J.P. & Co. (1972 SC): The Court emphasised that the overall impression of a mark, including its visual, phonetic, and conceptual elements, determines similarity. Minor differences are insufficient to dispel consumer confusion."),
  body("Toyota Jidosha Kabushiki Kaisha v. Prius Auto Industries (2017 SC): The Supreme Court addressed the territorial nature of trademark rights and likelihood of confusion in the context of a globally famous mark ('PRIUS') that had not been extensively used in India. The Court emphasised that likelihood of confusion must be assessed in the Indian market context."),
  subHeading("6. Defenses to Infringement"),
  body("Section 30 provides statutory defenses including: bona fide descriptive use; use indicating intended purpose; exhaustion of rights following lawful sale of goods; and use of one's own registered mark. These defenses operate as exceptions to the scope of the registered trademark owner's exclusive rights."),
  subHeading("7. Conclusion"),
  body("Likelihood of confusion is both the central test and the primary battleground in trademark infringement litigation. Indian courts have developed a nuanced, multi-factor framework for evaluating confusion that balances the interests of trademark proprietors, competitors, and consumers. The standard — the average consumer with imperfect recollection — ensures that trademark protection remains meaningful without unduly restricting fair competition."),
  spacer()
);

// B-Q6
children.push(
  qHeading(6, 10, "What is passing off? Distinguish it from trademark infringement. Discuss the three elements of a passing off action with reference to case law."),
  ansLabel(),
  subHeading("1. Introduction"),
  body("Passing off is an ancient common law tort that predates formal trademark registration systems. It protects traders from the misappropriation of their goodwill through misrepresentation by competitors. Unlike trademark infringement, which requires a registered mark, passing off provides relief to traders based on the reputation and goodwill they have built in their mark, get-up, trade name, or other business identifier — irrespective of formal registration. In India, passing off is expressly preserved by Section 27(2) of the Trade Marks Act, 1999, which makes it clear that registration is not necessary to maintain a passing off action."),
  subHeading("2. Definition and Nature of Passing Off"),
  body("Passing off occurs when a trader makes a misrepresentation to the public that his goods or services are those of another trader, thereby causing or likely to cause damage to that other trader's goodwill. The wrong lies not in using another's mark per se, but in representing one's goods as those of the plaintiff — 'passing off one's goods as the goods of another'. The foundation of the tort is the principle that no person shall be permitted to use a trade name or mark so as to deceive the public into thinking his goods are the goods of another trader."),
  subHeading("3. The Classical Trinity — Three Essential Elements"),
  body("The House of Lords in Reckitt & Colman Products Ltd. v. Borden Inc. (1990) famously articulated the three essential elements of a passing off action — goodwill, misrepresentation, and damage — a formulation universally adopted by Indian courts."),
  subHeading("(a) Goodwill"),
  body("The plaintiff must possess goodwill and reputation in the market. Goodwill, in this context, is the attractive force that brings in custom — it is the value associated with the brand or business identifier such that consumers recognise and purchase goods based on the mark. The plaintiff does not need a registered trademark; mere proof of consistent commercial use and consumer recognition suffices. Goodwill is a local concept — it must exist in the jurisdiction where the passing off occurs."),
  body("Case Law: In Hindustan Pencils v. India Stationery Products (1990), the plaintiff established goodwill in its 'NATARAJ' trademark for pencils and successfully obtained an injunction against the defendant's use of a similar mark. In Daimler Benz AG v. Hybo Hindustan (1994), the Delhi High Court recognized the goodwill of the 'BENZ' mark even though the plaintiff had no direct commercial presence in India at the time, acknowledging spill-over reputation."),
  subHeading("(b) Misrepresentation"),
  body("There must be a misrepresentation by the defendant — whether deliberate or innocent — that is likely to deceive a substantial portion of the relevant public into thinking that the defendant's goods or services are those of the plaintiff, or are connected with or endorsed by the plaintiff. Misrepresentation can arise from use of an identical or similar trade name, get-up, packaging, logo, domain name, or any other feature associated with the plaintiff's business."),
  body("Case Law: In Reckitt & Colman v. Borden (1990), the use of a look-alike lemon-shaped container for lemon juice was held to constitute misrepresentation. In Perry v. Truefitt (1842) — one of the earliest passing off cases — the defendant's use of a similar product name constituted misrepresentation. Indian courts in Pidilite Industries v. S.M. Associates (2003) held that the packaging of the defendant's adhesive product was deceptively similar to FEVICOL's get-up, constituting misrepresentation."),
  subHeading("(c) Damage"),
  body("The plaintiff must show that actual damage has resulted or is likely to result from the misrepresentation. Damage can take the form of: diversion of trade (consumers buying the defendant's product thinking it is the plaintiff's); loss of licensing or royalty income; dilution of the exclusivity of the plaintiff's mark; or harm to the reputation and goodwill of the plaintiff's business. The damage need not be proven precisely; it is sufficient to show that it is a reasonable and foreseeable consequence of the misrepresentation."),
  body("Case Law: In Reckitt & Colman v. Borden, the court noted that misrepresentation resulting in substitution of the plaintiff's product — thereby diverting sales — constitutes actionable damage. In Erven Warnink BV v. J. Townend & Sons (1979) ('Advocaat' case), the House of Lords recognized damage to the shared reputation of all legitimate producers of Advocaat as a valid head of damage in the extended form of passing off."),
  subHeading("4. Passing Off vs. Trademark Infringement — Key Distinctions"),
  body("Registration Requirement: Trademark infringement requires a registered mark. Passing off requires only goodwill and reputation through use; no registration is needed."),
  body("Nature of Right: Infringement is a statutory right under the Trade Marks Act, 1999. Passing off is a common law/equitable tort."),
  body("Scope: Infringement is strictly limited to goods/services covered by the registration (except for well-known marks). Passing off is more flexible — it extends to any mark, get-up, trade dress, or business identifier in which goodwill subsists."),
  body("Limitation Period: Infringement actions are governed by the Limitation Act. Passing off actions, being equitable in nature, are subject to the doctrine of laches — undue delay can defeat a claim."),
  body("Proof Required: In infringement, once the mark and use are established, confusion is presumed in many cases. In passing off, the plaintiff must affirmatively prove goodwill, misrepresentation, and damage."),
  subHeading("5. Extended Passing Off"),
  body("The doctrine of passing off has been extended beyond simple two-party cases to 'extended passing off' where a group of traders share a common reputation in a product name or description. In Erven Warnink v. Townend (Advocaat case), the House of Lords held that any trader within the legitimate class can sue a third party who misappropriates the product description, thereby damaging the collective reputation."),
  subHeading("6. Conclusion"),
  body("Passing off remains an indispensable remedy in trademark law, especially for businesses with unregistered marks, trade names, or distinctive get-up. Its three-element framework — goodwill, misrepresentation, and damage — provides a flexible and equitable basis for protecting commercial reputation. Indian courts have applied and developed this framework extensively, making passing off a robust complement to the statutory protection under the Trade Marks Act, 1999."),
  spacer()
);

// B-Q7
children.push(
  qHeading(7, 10, "Critically analyze the legal framework governing assignment of trademarks in India. Distinguish between assignment with and without goodwill and discuss situations where an assignment may be refused under the Act."),
  ansLabel(),
  subHeading("1. Introduction"),
  body("The assignment of trademarks — the permanent transfer of ownership of a trademark from one party (assignor) to another (assignee) — is an important commercial transaction that allows trademark rights to be bought, sold, and consolidated as business assets. The Trade Marks Act, 1999 provides a comprehensive framework for trademark assignment in Sections 37 to 45, balancing the need for commercial flexibility with the public interest in preventing consumer confusion and deception."),
  subHeading("2. Types of Trademark Assignment"),
  subHeading("(a) Assignment with Goodwill (Section 37)"),
  body("An assignment 'with goodwill' involves the transfer of not only the trademark itself but also the goodwill — the brand value, reputation, customer base, and business associations — that has accrued to the mark. In such a case, the assignee steps into the shoes of the assignor and acquires the full commercial value of the mark. The assignee can use the mark in all the goods or services for which the assignor had registered it. This is the most complete form of assignment and is commercially the most valuable."),
  subHeading("(b) Assignment without Goodwill (Section 37 — Gross Assignment)"),
  body("An assignment 'without goodwill' (also called a 'gross assignment') involves transfer of the mark itself but not the goodwill associated with it in the assignor's business. The assignee gets the registration rights but not the commercial reputation of the assignor in the mark. Historically, this type of assignment was viewed with suspicion because it might mislead consumers who associate the mark with the original owner's quality standards. However, the 1999 Act permits such assignment subject to conditions under Section 42 — particularly restrictions on simultaneous use that would be confusing to the public."),
  subHeading("3. Legal Requirements for Assignment"),
  body("Section 37 provides that the proprietor of a registered trademark may assign the trademark with or without the goodwill of the business concerned. Key requirements include:"),
  bullet("Assignment must be in writing and signed by or on behalf of the assignor (Section 37)."),
  bullet("The assignee must apply to the Registrar for registration of the assignment within three months of the date of the assignment instrument, or within such extended time as the Registrar may permit (Section 45)."),
  bullet("Form TM-P is used for assignment applications."),
  bullet("Where title has been acquired by transmission (by operation of law or by court order), evidence of such transmission must be submitted."),
  subHeading("4. Assignment of Unregistered Trademarks"),
  body("Section 39 allows assignment of unregistered trademarks, but only along with the goodwill of the business in which the mark is used. This restriction prevents trafficking in unregistered marks stripped of their business context. An assignment of an unregistered mark without goodwill is not valid under the Act."),
  subHeading("5. Situations Where Assignment May Be Refused"),
  body("The Registrar or courts may refuse or restrict an assignment in the following situations:"),
  subHeading("(a) Assignment Creating Confusion — Section 40"),
  body("The Act prohibits an assignment that would result in the creation of exclusive rights in two or more persons in relation to the same or similar goods or services, in different parts of India, where such concurrent use would be likely to cause confusion or deception among the public. For example, if X assigns his mark to Y for Northern India and retains the mark for Southern India, and both use the mark for the same goods, the resulting situation may cause consumer confusion. The Registrar can refuse to register such assignments."),
  subHeading("(b) Assignment Resulting in Multiple Exclusive Rights in Same Goods — Section 41"),
  body("Where an assignment would result in different persons holding registered trademarks that are identical or nearly resembling each other in respect of the same goods or services, and each person claims exclusive rights, the assignment may be refused as it would lead to public confusion."),
  subHeading("(c) Assignment without Goodwill — Concurrent Use Issues (Section 42)"),
  body("Where a mark is assigned without goodwill, both the assignor (who retains the goodwill) and the assignee (who has the mark) may simultaneously use the mark in the market, leading to confusion. To address this, Section 42 allows the Registrar to restrict or condition the concurrent use of the mark by requiring the assignee to use the mark in a way that distinguishes its goods from those of the assignor."),
  subHeading("6. Critical Analysis"),
  body("The framework for trademark assignment under the 1999 Act is generally well-structured but has attracted some criticism. First, the requirement of recording assignment within three months (Section 45) creates uncertainty during the interim period — the assignee has rights in equity but no legal recognition until recorded. Second, the provisions on assignment without goodwill, while commercially useful, can undermine the consumer protection function of trademarks if quality control is lost. Third, cross-border trademark assignments in the context of multinational corporations raise complex questions about which national law governs and whether Indian restrictions apply to globally integrated trademark portfolios."),
  subHeading("7. Conclusion"),
  body("The legal framework governing trademark assignment in India strikes a reasonable balance between commercial flexibility and consumer protection. The distinction between assignment with and without goodwill recognises commercial realities while guarding against consumer deception. The provisions restricting assignments that create confusion reflect the fundamental principle that trademarks are indicators of source and must not be allowed to mislead the public."),
  spacer()
);

// B-Q8
children.push(
  qHeading(8, 10, "Discuss the defenses available to a defendant in a trademark infringement action under the Trade Marks Act, 1999. Critically evaluate the defense of 'descriptive use' and 'exhaustion' with reference to judicial decisions."),
  ansLabel(),
  subHeading("1. Introduction"),
  body("While the Trade Marks Act, 1999 grants registered trademark owners exclusive rights and strong remedies for infringement, it also recognizes that these rights must be balanced against the interests of competitors, traders, and the public. Section 30 of the Act enumerates specific defenses to trademark infringement, providing a 'savings clause' that carves out uses which do not constitute infringement even when a registered mark is involved. Among these, the defenses of 'descriptive use' (fair use) and 'exhaustion' (first sale doctrine) are the most frequently litigated and analytically important."),
  subHeading("2. Defenses Available Under Section 30"),
  subHeading("(a) Bona Fide Descriptive Use — Section 30(2)(a)"),
  body("A person does not infringe a registered trademark if they use the mark in a bona fide manner to describe the kind, quality, quantity, intended purpose, value, geographical origin, or other characteristics of the goods or services. The key conditions for this defense are: (i) the use must be descriptive in nature; (ii) the use must be honest and in good faith; and (iii) the use must be in accordance with honest practices in industrial and commercial matters."),
  subHeading("(b) Use Indicating Intended Purpose — Section 30(2)(b)"),
  body("Using a registered trademark to indicate the intended purpose of a product or service (for example, indicating compatibility with a branded product) is permissible, provided the use is honest and necessary to indicate the intended purpose. This is particularly relevant for spare parts, accessories, and software compatibility disclosures. Example: A car parts manufacturer can lawfully state 'Compatible with MARUTI SUZUKI vehicles' without infringing the MARUTI SUZUKI trademark, as this indicates intended purpose, not origin."),
  subHeading("(c) Exhaustion of Rights — Section 30(3)"),
  body("Under the doctrine of exhaustion, a registered trademark owner's rights in respect of goods bearing the trademark are exhausted once those goods are lawfully put on the Indian market by the proprietor or with the proprietor's consent. Subsequent dealers, resellers, or distributors who trade in those genuine goods cannot be held liable for trademark infringement, as the proprietor has already derived benefit from the first sale."),
  subHeading("(d) Use of Proprietor's Own Registered Mark"),
  body("A person who holds their own registered trademark cannot be sued for infringement by another trademark owner, at least for the goods or services for which their mark is registered, within the scope of their own registration."),
  subHeading("(e) Prior Use Defense (Section 34)"),
  body("A person who has been using a mark (or a mark nearly resembling the registered mark) continuously since a date prior to the date of registration of the registered mark is entitled to continue such use. This 'prior use' defense is available even against a registered trademark holder."),
  subHeading("(f) Other Defenses"),
  body("Additional defenses include: acquiescence (Section 33) — where the registered proprietor has acquiesced in the use of a later registered trademark for a continuous period of five years with knowledge, they lose the right to challenge that later registration or its use; delay and laches; comparative advertising; and honest concurrent use."),
  subHeading("3. Critical Evaluation of Descriptive Use Defense"),
  body("The descriptive use defense under Section 30(2)(a) recognizes that descriptive words must remain available to all traders so that they can communicate truthfully about their products. The defense prevents trademark owners from monopolizing everyday language."),
  body("However, courts have been vigilant in ensuring that the 'descriptive use' defense is not used as a cloak for parasitic use of a well-known mark. In Reckitt Benckiser India Ltd. v. Wyeth Ltd. (2013), the Delhi High Court held that a defendant claiming descriptive use must demonstrate that the use is genuinely descriptive of the product's qualities and not a use of the mark as a badge of origin or an attempt to free-ride on the plaintiff's reputation."),
  body("The defense also requires honesty — use that is technically descriptive but commercially opportunistic may be denied. In Marico Ltd. v. Agro Tech Foods Ltd. (2010), the Delhi High Court held that even comparative advertising using a competitor's trademark must be truthful and not misleading. Use that exaggerates or distorts facts cannot claim the descriptive/fair use defense."),
  subHeading("4. Critical Evaluation of Exhaustion Defense"),
  body("The exhaustion doctrine prevents trademark owners from using their marks to partition markets and control resale after goods have lawfully entered commerce. India follows the principle of international exhaustion, meaning that once the goods are placed on the market anywhere in the world by the proprietor or with consent, rights over those specific goods are exhausted. This was affirmed in Samsung Electronics Co. Ltd. v. Kapil Wadhwa (2012), where the Delhi High Court (Division Bench) held that India follows international exhaustion — Samsung could not prevent a parallel importer from importing and selling genuine Samsung products in India without authorization."),
  body("However, the exhaustion defense has important limitations. Under Section 30(4), the defense is unavailable if the proprietor has legitimate reasons to oppose further commercialisation, especially where the condition of the goods has been altered or impaired after they were put on the market. In cases where goods are repackaged, relabelled, or modified in any way, the exhaustion defense does not apply."),
  body("The Bata India Ltd. v. Pyare Lal (1985) case established that resale of genuine branded goods does not constitute infringement, laying the foundation for the Indian exhaustion doctrine. Courts have also held that grey market goods — genuine goods imported through parallel channels without the manufacturer's Indian authorization — are protected by the exhaustion doctrine, as long as the goods are authentic and unmodified."),
  subHeading("5. Conclusion"),
  body("The defenses available under Section 30 of the Trade Marks Act, 1999 play a crucial role in preventing trademark owners from over-reaching and stifling legitimate competition. The descriptive use defense preserves the public domain of common language, while the exhaustion defense promotes free trade and consumer access to genuine goods. Indian courts have applied both defenses with a careful eye on commercial realities, ensuring that trademarks serve their legitimate function as indicators of origin without becoming tools of market monopolisation."),
  spacer()
);

// B-Q9
children.push(
  qHeading(9, 10, "Discuss the law relating to geographical indications in India under the Geographical Indications of Goods (Registration and Protection) Act, 1999. How do GIs interact with trademark law — can they coexist or does one override the other?"),
  ansLabel(),
  subHeading("1. Introduction"),
  body("Geographical Indications (GIs) represent a unique form of intellectual property that bridges heritage, geography, and commerce. They protect the collective reputation of products that owe their distinctive qualities to their place of origin. India enacted the Geographical Indications of Goods (Registration and Protection) Act, 1999 (the GI Act) to comply with TRIPS obligations (Article 22–24) and to protect India's rich tradition of regionally distinctive products. The interaction between GI law and trademark law is a fascinating area where collective rights meet individual commercial rights."),
  subHeading("2. What is a Geographical Indication?"),
  body("Under Section 2(1)(e) of the GI Act, a geographical indication is an indication which identifies such goods as agricultural goods, natural goods, or manufactured goods as originating from a territory, region, or locality in that territory where a given quality, reputation, or other characteristic of such goods is essentially attributable to its geographical origin."),
  body("Examples of registered GIs in India include: Darjeeling Tea (the first Indian GI, registered in 2004), Basmati Rice, Alphonso Mangoes, Kancheepuram Silk, Mysore Silk, Chanderi Fabric, Tirupathi Laddu, Kolhapuri Chappals, Nagpur Orange, and Pochampally Ikat."),
  subHeading("3. Registration Process Under the GI Act"),
  body("Any association of persons, producers, organization, or authority established by law representing the interest of producers of the goods can apply for GI registration to the Geographical Indications Registry (located in Chennai). The application must detail: the geographical area, the nature of the goods, the description of quality or characteristics attributable to geographical origin, and the standards applicable. The GI is registered for a period of 10 years and is renewable indefinitely. Once registered, producers within the designated geographical area who comply with the specified standards can use the GI."),
  subHeading("4. Authorised Users"),
  body("In addition to registration of the GI, individual producers within the region can be registered as 'authorised users' under Section 17 of the GI Act. Only registered proprietors and authorised users can use the GI mark on products. The GI Registry maintains separate records for the GI itself and its authorised users."),
  subHeading("5. Protection and Infringement Under the GI Act"),
  body("Once registered, a GI is protected against: use by persons who are not authorised users; use of false geographical indications; use of the GI in a manner that misleads consumers about the origin of the goods; and use of the GI for goods not originating from the specified region. The remedies include civil remedies (injunction, damages, account of profits) and criminal penalties. The GI Act follows a strict liability approach — even innocent use of a false GI is actionable."),
  subHeading("6. GIs and Trademark Law — Interaction and Coexistence"),
  body("The relationship between GIs and trademarks is complex and governed by principles of mutual exclusivity in certain cases and coexistence in others."),
  subHeading("(a) GI Cannot Be Registered as a Trademark"),
  body("Section 25 of the Trade Marks Act, 1999 (and Section 2(1)(zb) read with Section 9) prohibits the registration of a trademark that consists of or comprises a false indication of geographical origin, or that is deceptively similar to a registered geographical indication for goods of the same class. Thus, no individual trader can register a trademark that incorporates a GI, preventing monopolisation of a collective identifier."),
  subHeading("(b) Pre-Existing Trademarks and GIs"),
  body("A significant conflict arises where an individual trademark was registered before the GI registration. Section 24 of the GI Act provides a partial solution: where a trademark was applied for or registered in good faith before the GI was registered, the trademark registration is not invalidated solely by reason of the GI registration. The trademark continues to be valid for the goods for which it was registered, even if it coincides with a GI."),
  subHeading("(c) Well-Known Marks vs. GIs"),
  body("Courts have had to balance the rights of well-known trademark owners against GI rights. In the 'BASMATI' dispute between India and Pakistan at the international level, and in various Indian trademark disputes involving 'CHAMPAGNE' and 'SCOTCH WHISKY', courts have emphasized that GIs represent collective regional rights that ordinarily override individual trademark claims for the same goods."),
  subHeading("(d) Certification Marks and GIs"),
  body("The Trade Marks Act, 1999 provides for certification marks (Chapter VIII), which in some ways overlap with GIs — both can certify origin and quality. However, certification marks are registered by an individual or body that sets standards, whereas GIs are collective rights of all producers in a region. In India, both can coexist, though registration of the same indicia as both a certification mark and a GI is avoided to prevent confusion."),
  subHeading("7. Critical Evaluation"),
  body("India's GI framework has been criticized for limited international recognition (GI registration in India does not automatically provide protection in other countries), the slow pace of registration and enforcement, lack of awareness among artisans and producers about GI rights, and the absence of a strong market premium for most registered GIs. On the positive side, GIs have provided legal ammunition against biopiracy and have helped promote India's traditional and artisanal products globally."),
  subHeading("8. Conclusion"),
  body("GIs and trademarks serve complementary but distinct purposes in the intellectual property ecosystem. Trademarks identify the commercial source of goods from a single enterprise; GIs identify goods whose qualities are tied to a collective regional heritage. Indian law, through the GI Act 1999 and the Trade Marks Act 1999, provides for a nuanced coexistence — generally preventing GIs from being appropriated as individual trademarks while protecting pre-existing trademark rights. The two regimes can coexist where they cover different goods or where historical rights are preserved."),
  spacer()
);

// B-Q10
children.push(
  qHeading(10, 10, "Analyze the adequacy of civil and criminal remedies available for trademark infringement under Indian law. Do you think Anton Piller orders and John Doe orders have strengthened enforcement? Justify."),
  ansLabel(),
  subHeading("1. Introduction"),
  body("Effective enforcement of trademark rights depends not merely on substantive rights but also on the adequacy of the remedial framework. The Trade Marks Act, 1999 provides a range of civil and criminal remedies, supplemented by equitable orders developed by Indian courts. The adequacy of these remedies in the digital and global commercial environment — particularly the innovative enforcement tools of Anton Piller orders and John Doe orders — is a subject of increasing judicial and academic attention."),
  subHeading("2. Civil Remedies"),
  body("Section 135 of the Trade Marks Act, 1999 provides the following civil remedies to an aggrieved trademark owner:"),
  bullet("Injunction: An injunction restraining the defendant from continuing the infringing activity is the primary and most sought-after civil remedy. Courts can grant temporary injunctions (interlocutory orders) pending final decision, as well as permanent injunctions after trial. The classic test for a temporary injunction (from American Cyanamid Co. v. Ethicon Ltd., adopted by Indian courts) is: (i) prima facie case; (ii) balance of convenience; and (iii) irreparable harm. An interlocutory injunction can be a highly effective enforcement tool as it immediately stops infringing conduct."),
  bullet("Damages or Account of Profits: The plaintiff can claim either actual damages suffered due to infringement or an account of profits earned by the defendant from the infringing use. The court awards whichever is appropriate based on the evidence. Indian courts have increasingly awarded punitive or enhanced damages in cases of flagrant or deliberate infringement."),
  bullet("Delivery Up and Destruction: The court can order the defendant to deliver up all infringing goods, labels, packaging, and materials to the plaintiff or to a court officer. Infringing goods may be ordered to be destroyed. This prevents infringing goods from re-entering the market through back channels."),
  bullet("Enhanced Damages in Cases of Fraud: In cases where the infringement was deliberate or fraudulent, courts can award enhanced damages as a deterrent. The Supreme Court in Hindustan Lever Ltd. v. Nirma Ltd. recognized the principle of enhanced damages in flagrant infringement cases."),
  subHeading("3. Criminal Remedies"),
  body("Section 103 of the Trade Marks Act, 1999 provides for criminal punishment for applying false trademarks or trade descriptions, with imprisonment of up to three years and/or a fine of up to two lakh rupees, with enhanced penalties for second and subsequent offences. Section 104 deals with the selling of goods with false marks. Criminal prosecution, while available, is relatively infrequently pursued by trademark owners who typically prefer the faster civil remedy of injunction. However, criminal complaints are effective in cases of large-scale counterfeiting operations as they lead to raids, seizures, and imprisonment of infringers."),
  subHeading("4. Border Measures"),
  body("The Customs (Import of Goods at Concessional Rate of Duty) Rules, 2017 and the Intellectual Property Rights (Imported Goods) Enforcement Rules, 2007 allow trademark owners to register their marks with Customs authorities. Customs officers can then seize counterfeit or infringing goods at the border before they enter the Indian market — a proactive enforcement mechanism particularly important for luxury goods and electronics."),
  subHeading("5. Anton Piller Orders (Search Orders)"),
  body("An Anton Piller order (now called a 'Search Order' in England under the Civil Procedure Rules) is an ex parte court order that authorises the plaintiff (or their solicitor) to enter the defendant's premises to search for, inspect, and seize infringing goods or documents without prior notice to the defendant. It is named after the English case Anton Piller KG v. Manufacturing Processes Ltd. (1976). The order is particularly valuable where there is a real risk that the defendant will destroy evidence if given advance notice."),
  body("Indian courts have extensively used Anton Piller-style orders in trademark and copyright infringement cases, especially in the context of piracy and counterfeiting. In Autodesk Inc. v. AVT Shankardass (2010), the Delhi High Court granted a local commissioner-based inspection order akin to an Anton Piller order in a software piracy matter. These orders have strengthened enforcement by: enabling surprise inspections of suspected counterfeit goods warehouses; preventing destruction of evidence; enabling identification of the scale of infringing operations; and facilitating the filing of comprehensive suits after evidence is secured."),
  subHeading("6. John Doe Orders (Ashok Kumar Orders)"),
  body("A John Doe order (called an 'Ashok Kumar order' in India, after the case Taj Television Ltd. v. Rajan Mandal) is an ex parte injunction obtained against unnamed or unknown defendants. It is granted when the plaintiff cannot identify all infringers at the time of filing the suit. The order typically restrains 'John Doe, Ashok Kumar, and all others acting in concert' from infringing the plaintiff's trademark or copyright."),
  body("John Doe orders have been particularly transformative in the digital age. In the entertainment and trademark context, such orders have been obtained against unknown websites, street vendors, counterfeiters, and internet service providers who host or distribute infringing content. In Reliance Big Entertainment v. Multivision Network (2012) and subsequent film piracy matters, courts issued John Doe orders against cable operators and websites distributing infringing content. In the trademark context, John Doe orders have been obtained to restrain unknown counterfeit vendors from selling fake luxury goods."),
  body("These orders have strengthened enforcement by: enabling action against large-scale unidentified infringers (counterfeit street vendors at markets, online sellers); providing law enforcement agencies with judicial sanction for raids; serving as deterrents to potential infringers; and allowing right holders to take swift action at scale."),
  subHeading("7. Critical Assessment of Adequacy"),
  body("While the civil and criminal remedies and innovative equitable orders under Indian trademark law are comprehensive, there are significant implementation challenges. Litigation is protracted — even interim injunction proceedings can take months. The quantum of damages awarded in most cases is considered insufficient to deter determined counterfeiters. Criminal prosecutions are rarely pursued to conviction. Digital counterfeiting — through e-commerce platforms and social media — presents new challenges not fully addressed by current law. The 1999 Act predates the rise of e-commerce, and courts have had to adapt existing remedies creatively."),
  subHeading("8. Conclusion"),
  body("Anton Piller orders and John Doe orders have undoubtedly strengthened trademark enforcement in India by providing right holders with powerful tools to act swiftly against infringers — both known and unknown — and to preserve evidence before it can be destroyed. The civil remedies framework under the 1999 Act is generally adequate in scope, though the pace of judicial proceedings and the quantum of damages awarded remain areas for improvement. A combination of robust statutory remedies, innovative equitable orders, and administrative measures (like border controls) collectively provide a reasonably comprehensive enforcement toolkit for Indian trademark owners."),
  spacer()
);

// B-Q11
children.push(
  qHeading(11, 10, "'Domain name disputes are trademark disputes in disguise.' Critically examine this statement in light of ICANN's UDRP policy and Indian jurisprudence on domain names, referring to leading cases."),
  ansLabel(),
  subHeading("1. Introduction"),
  body("The proposition that 'domain name disputes are trademark disputes in disguise' reflects the deep functional overlap between domain names and trademarks in the digital economy. A domain name — the unique address by which a website is identified on the internet (e.g., www.google.com) — can serve as a powerful brand identifier, just as a trademark does in the physical marketplace. The phenomena of 'cybersquatting' (registering domain names identical or similar to trademarks in bad faith) and 'typosquatting' (registering common misspellings) have forced trademark law to expand into the digital domain, leading to the development of specialized dispute resolution mechanisms."),
  subHeading("2. The Trademark-Domain Name Nexus"),
  body("A domain name serves not merely as a technical address but as a commercial identifier. Consumers use domain names to find businesses online, often assuming that a company's website is at 'www.[trademark].com'. When a third party registers a domain name identical or similar to a well-known trademark, this can: divert consumer traffic; damage the trademark owner's online brand presence; enable phishing and fraud; and cause confusion about the official source of goods or services. These are precisely the harms that trademark law seeks to prevent — hence the characterization of domain name disputes as trademark disputes."),
  subHeading("3. ICANN's Uniform Domain Name Dispute Resolution Policy (UDRP)"),
  body("ICANN (Internet Corporation for Assigned Names and Numbers) adopted the UDRP in 1999 as a fast, cost-effective, and global mechanism for resolving domain name disputes involving trademark rights. The UDRP applies to all generic top-level domains (gTLDs) such as .com, .net, .org."),
  subHeading("Three-Part Test Under UDRP:"),
  body("A complainant must prove all three of the following elements: (1) The domain name is identical or confusingly similar to a trademark or service mark in which the complainant has rights; (2) The registrant has no rights or legitimate interests in respect of the domain name; (3) The domain name has been registered and is being used in bad faith."),
  subHeading("Indicators of Bad Faith Under UDRP (Paragraph 4(b)):"),
  bullet("Registered to sell to the trademark owner at a profit exceeding out-of-pocket costs."),
  bullet("Registered to prevent the trademark owner from reflecting its mark in a domain name, as part of a pattern of such conduct."),
  bullet("Registered to disrupt the business of a competitor."),
  bullet("Used to intentionally attract internet users for commercial gain by creating confusion with the complainant's mark."),
  body("UDRP proceedings are conducted online before accredited dispute resolution providers — primarily WIPO Arbitration and Mediation Center and Forum (formerly NAF). Remedies are limited to cancellation or transfer of the domain; no monetary damages are available. Proceedings typically conclude within 60 days."),
  subHeading("4. Leading UDRP Cases"),
  body("In Telstra Corporation Ltd. v. Nuclear Marshmallows (WIPO, 2000), it was held that passive holding of a domain name (without active use) can constitute bad faith if the complainant's mark is well-known and there is no conceivable legitimate use by the respondent. In Julia Roberts v. Russell Boyd (WIPO, 2000), the domain 'juliaroberts.com' was transferred, recognizing celebrity name rights under the UDRP. In Satyam Infoway Ltd. v. Sifynet Solutions (India SC, 2004), the Supreme Court held that internet domain names are subject to the same legal norms as trademarks and that the tort of passing off applies to domain name misuse."),
  subHeading("5. Indian Jurisprudence on Domain Names"),
  body("India does not have a dedicated domain name law; disputes are resolved through UDRP (for gTLDs), the National Internet Exchange of India's '.IN Dispute Resolution Policy' (INDRP) for .in domains, and through civil courts applying trademark and passing off principles."),
  body("Satyam Infoway Ltd. v. Sifynet Solutions Pvt. Ltd. (2004 SC): The Supreme Court of India definitively held that domain names serve as business identifiers and are protectable under passing off principles. The Court held that domain names can be the subject of proprietary rights and that misrepresentation through a confusingly similar domain name is actionable. This case established that trademark law principles apply to domain name disputes in India."),
  body("Yahoo Inc. v. Akash Arora (1999, Delhi HC): The Delhi High Court granted an injunction against 'yahooindia.com', holding that the domain name was deceptively similar to Yahoo's well-known trademark and constituted passing off. This was one of India's earliest domain name decisions and applied trademark principles to cyberspace."),
  body("Rediff Communication Ltd. v. Cyberbooth (1999, Bombay HC): The Bombay High Court restrained the use of 'Rediff.com' by a cybercafe, holding that online domain names are akin to trade names or trademarks and attract similar legal protection."),
  subHeading("6. Critical Examination of the Statement"),
  body("The statement that 'domain name disputes are trademark disputes in disguise' is largely accurate but requires nuance. On one hand, the UDRP itself is fundamentally predicated on trademark rights — without a trademark, a complainant has no standing under UDRP. The tests applied (likelihood of confusion, bad faith, legitimate interests) are essentially trademark law concepts applied in the digital context."),
  body("On the other hand, domain names have characteristics that distinguish them from trademarks. Unlike trademarks, domain names are unique — only one registrant can hold 'www.apple.com'. Domain names are globally visible and cannot be territorially limited as trademarks can. Domain names can serve purely functional purposes (e.g., internal corporate websites) with no trademark significance. Personal names, geographic names, and descriptive terms can be registered as domain names without trademark protection."),
  body("Furthermore, cyberpiracy and cybersquatting — registering domain names in bad faith to sell to trademark owners — is a wrong that goes beyond trademark infringement in the traditional sense. The UDRP addresses this specifically, creating a sui generis remedy that incorporates trademark principles but is not pure trademark law."),
  subHeading("7. Conclusion"),
  body("Domain name disputes are substantially, though not entirely, trademark disputes. The UDRP's three-element test is rooted in trademark concepts, and Indian courts have consistently applied trademark and passing off principles to resolve domain name conflicts. However, the unique characteristics of domain names — their global reach, functional duality, and technical uniqueness — mean that domain name law has evolved as a specialized hybrid discipline that borrows heavily from but is not wholly subsumed within trademark law."),
  spacer()
);

// B-Q12
children.push(
  qHeading(12, 10, "Evaluate the legal protection available for publicity rights in India. Is there an independent statutory framework for publicity rights or does protection come through allied laws like passing off and trademarks? Suggest legislative reforms."),
  ansLabel(),
  subHeading("1. Introduction"),
  body("The right of publicity — the right of individuals (most commonly celebrities) to control the commercial use of their name, image, likeness, voice, or persona — represents an intersection of personality rights, privacy law, and intellectual property law. As celebrity-driven commerce has exploded in the digital age, the unauthorized commercial exploitation of a person's identity without consent has become a pressing legal issue. India, unlike the United States (which has state-level right of publicity statutes) or the United Kingdom (which relies on the tort of passing off and the data protection framework), has no dedicated statute for publicity rights. Protection in India comes through a patchwork of constitutional guarantees, judicial pronouncements, and allied IP laws."),
  subHeading("2. Nature and Scope of Publicity Rights"),
  body("Publicity rights encompass: the right to control use of one's name in commercial advertising; the right to prevent unauthorized use of one's image or photograph in endorsements; the right to prevent others from commercially exploiting one's likeness, voice, signature, or persona; and the right to benefit financially from commercial exploitation of one's celebrity persona. The right is sometimes divided into (a) the 'right of publicity' — the economic right to exploit one's persona, and (b) the 'right of privacy/personality' — the moral right to control the use of one's image."),
  subHeading("3. Constitutional Basis"),
  body("Indian courts have increasingly derived publicity rights from the fundamental right to privacy under Article 21 of the Constitution, which was expansively interpreted by the Supreme Court in Justice K.S. Puttaswamy v. Union of India (2017). The nine-judge bench unanimously held that privacy includes the right to control information about oneself and that personality rights are a component of the constitutional right to privacy. This provides a constitutional foundation for publicity rights even in the absence of a dedicated statute."),
  subHeading("4. Judicial Recognition — Common Law and Equitable Remedies"),
  body("(a) Passing Off: Indian courts have used the tort of passing off to protect celebrities against unauthorized commercial exploitation of their name and image. In Titan Industries Ltd. v. Ramkumar Jewellers (2012), the Delhi High Court held that a celebrity's image rights are a form of goodwill and that unauthorized use of their image in advertising constitutes passing off. The court recognized that a celebrity's name and image have commercial value, and misappropriation creates a false impression of endorsement."),
  body("(b) ICC Development (International) Ltd. v. Arvee Enterprises & Anr. (2003, Delhi HC): The Delhi High Court held that the right of publicity vests in the individual and includes the right to control the commercial use of one's personality, name, and likeness. The court recognized a distinct 'right of publicity' as an enforceable right in India, actionable through civil suits for damages and injunctions."),
  body("(c) Shivaji Rao Gaikwad (Rajinikanth) v. Varsha Productions (2015, Madras HC): The Madras High Court protected the superstar's likeness and character from being used in a derogatory film, recognizing personality rights as actionable in India."),
  body("(d) Trademark Law: Celebrities can register their names or signatures as trademarks under the Trade Marks Act, 1999 (as service marks or in merchandising classes). Unauthorized use of a registered celebrity trademark is directly actionable as trademark infringement. However, this protection is limited to registered marks in specified classes and does not cover all forms of personality exploitation."),
  subHeading("5. Limitations of Current Legal Framework"),
  body("Despite judicial ingenuity, the current framework has significant gaps: Passing off requires proof of goodwill, misrepresentation, and damage — a high threshold. Trademark protection is limited to registered marks in specified classes and requires active enforcement. There is no protection for non-celebrity individuals whose image is misappropriated commercially. Post-mortem publicity rights (protecting the estate of deceased celebrities) are unrecognized in India, unlike in several US states. Social media and deepfake technologies have created new avenues for personality exploitation that existing law does not adequately address."),
  subHeading("6. Legislative Reforms Suggested"),
  body("India urgently needs a dedicated Right of Publicity statute. The following reforms are recommended:"),
  bullet("Enact a standalone Right of Publicity Act that recognizes the economic and moral components of publicity rights for all individuals (not just celebrities), without requiring proof of goodwill or misrepresentation."),
  bullet("Define the scope of protected attributes: name, photograph, likeness, voice, signature, biometric data, and AI-generated replicas of a person's persona."),
  bullet("Recognize post-mortem publicity rights extending for at least 50 years after death, allowing estates to commercially exploit and protect the deceased's persona."),
  bullet("Provide clear exceptions for news reporting, satire, commentary, artistic expression, and educational use to balance publicity rights against freedom of speech."),
  bullet("Introduce specific provisions against deepfakes and synthetic media that replicate a person's likeness or voice without consent."),
  bullet("Prescribe statutory damages to incentivise litigation without requiring detailed proof of financial loss."),
  subHeading("7. Conclusion"),
  body("The legal protection for publicity rights in India, while growing through judicial innovation, remains inadequate and fragmented. Protection comes primarily through constitutional privacy rights, the tort of passing off, and trademark law — none of which provides a complete or systematic remedy. A dedicated Right of Publicity statute, incorporating international best practices while accounting for India's constitutional framework, is both necessary and timely. Such legislation would provide clearer rights, more certain remedies, and better protection for individuals against an increasingly exploitative commercial media landscape."),
  spacer()
);

// B-Q13
children.push(
  qHeading(13, 10, "A multinational corporation holds a trademark registered globally. It seeks to license the trademark to a franchisee in India. Advise the corporation on the legal requirements, permissible limitations, and potential risks under the Trade Marks Act, 1999, including quality control obligations."),
  ansLabel(),
  subHeading("1. Introduction"),
  body("This is a practical problem-based question requiring application of trademark licensing law, franchise law, quality control obligations, and cross-border IP considerations under the Trade Marks Act, 1999. The multinational corporation (MNC) must navigate multiple layers of Indian law to ensure its trademark is effectively licensed, its brand integrity is maintained, and its legal position is protected."),
  subHeading("2. Preliminary Step: Ensure Trademark Registration in India"),
  body("Before licensing a trademark in India, the MNC must ensure that the trademark is registered in India under the Trade Marks Act, 1999. Global registration (through the Madrid Protocol or otherwise) does not automatically confer Indian registration rights. The MNC should: verify that the mark is on the Indian Trade Marks Register for the relevant goods/services classes; if not registered, file for registration and consider an interim licensing arrangement or rely on well-known mark protection under Section 11(2) during the pendency of registration; ensure there are no pending cancellation proceedings or adverse trademark registrations by third parties in India."),
  subHeading("3. Legal Requirements for a Valid Trademark License in India"),
  subHeading("(a) Registration of Licensee as Registered User (Section 48)"),
  body("The MNC (as proprietor) and the Indian franchisee must jointly apply to the Registrar of Trade Marks to register the franchisee as a 'registered user' of the mark in India (Form TM-U). The application must state: the goods or services in relation to which the mark will be used; the period of the license; the territory; the conditions and restrictions; and a statement of the relationship between proprietor and registered user. A registered user has rights to use the mark and has standing in infringement proceedings. An unregistered licensee can use the mark contractually but may lack standing to independently sue infringers."),
  subHeading("(b) Written License Agreement"),
  body("The MNC should execute a comprehensive written license (franchise) agreement with the Indian franchisee covering: scope of license (exclusive/non-exclusive, territory, goods/services); duration and renewal conditions; quality standards and compliance obligations; royalty/fee structure; MNC's quality control rights (inspection, audit, product testing); grounds for termination (breach of quality standards, insolvency, competitive infringement); non-compete and non-disclosure clauses; and dispute resolution mechanism (arbitration preferred, given cross-border nature)."),
  subHeading("(c) Foreign Exchange and Regulatory Compliance"),
  body("Royalty payments from India to the MNC are governed by the Foreign Exchange Management Act (FEMA), 1999 and RBI guidelines. Royalty remittances for trademark licensing are generally permissible under the automatic route up to permissible limits. The franchise agreement should comply with FEMA regulations, and the franchisee should obtain necessary regulatory approvals."),
  subHeading("4. Permissible Limitations Under the Trade Marks Act"),
  body("The MNC may impose the following permissible limitations in the license: Territorial restriction — limit the franchise to specific states or cities in India; Exclusive use — grant an exclusive license for certain products, preventing the MNC from directly competing in that territory; Quality specifications — mandate use of specific ingredients, raw materials, suppliers, or manufacturing processes consistent with global brand standards; Sublicensing prohibition — prevent the franchisee from sublicensing the mark without express written consent; Non-compete obligations — prohibit the franchisee from using competing marks during and for a defined period after the license term."),
  subHeading("5. Quality Control Obligations"),
  body("This is the most critical aspect of franchise trademark licensing. The Trade Marks Act, 1999 (Section 49 and the framework of Section 48) requires that the proprietor exercise control over the use of the mark by the licensee. The MNC must:"),
  bullet("Specify quality standards in the franchise agreement — including product specifications, service standards, use of packaging, advertising guidelines, and operational manuals."),
  bullet("Reserve rights of inspection and audit — the right to inspect the franchisee's premises, products, and records to verify compliance with brand standards."),
  bullet("Insist on mandatory training — ensure that the franchisee's staff are trained in brand standards."),
  bullet("Implement a quality monitoring system — conduct periodic product quality tests, mystery customer audits, and compliance reviews."),
  bullet("Provide for immediate termination on quality breach — include provisions for immediate suspension and termination of the license if quality standards are not met."),
  body("Failure to maintain quality control constitutes a 'naked license' which can result in: the mark losing its distinctiveness or becoming deceptive; cancellation of the registered trademark under Section 57 by the Registrar or court; and damage to the MNC's global brand reputation."),
  subHeading("6. Potential Risks"),
  bullet("Risk of Naked License: If the MNC fails to actively monitor quality, the trademark may be cancelled on the ground that it has become deceptive or generic — a risk especially acute in consumer goods and food sectors."),
  bullet("Risk of Mark Becoming Deceptive: If the franchisee uses the mark for sub-standard goods without oversight, the Indian public may associate the global mark with poor quality, damaging the brand even in markets where the MNC operates directly."),
  bullet("Risk of Parallel Imports: If the MNC licenses the mark to another franchisee in a neighbouring country, parallel imports of cheaper goods may enter India, potentially competing with the Indian franchisee's products under the same mark."),
  bullet("Risk of Infringement by Third Parties: The MNC must ensure the license agreement obligates the franchisee to promptly report and assist in actions against infringers of the licensed mark in India."),
  bullet("Risk of Franchisee's Insolvency: The license agreement should specify what happens to the trademark license in the event of the franchisee's insolvency — trademark rights should revert to the MNC immediately."),
  bullet("Risk of Franchisee Acquiring Rights: Long use by the franchisee without proper documentation could, in rare cases, give rise to claims of prior use rights by the franchisee, especially if the MNC's Indian registration lapses. Proper registration maintenance and documentation are essential."),
  subHeading("7. Conclusion"),
  body("The MNC seeking to license its trademark to an Indian franchisee must comply with the registered user provisions of the Trade Marks Act, 1999, execute a comprehensive franchise agreement, maintain rigorous quality control, and comply with India's foreign exchange regulations. The most critical obligation is quality control — both as a legal requirement under Indian trademark law and as a practical necessity for protecting the global brand's reputation and value in the Indian market."),
  spacer()
);

// B-Q14
children.push(
  qHeading(14, 10, "X has been using a mark for 15 years without registration. Y later registers the same mark and sues X for infringement. X claims prior use as a defense. Analyze X's legal position under the Trade Marks Act, 1999."),
  ansLabel(),
  subHeading("1. Introduction"),
  body("This problem raises one of the most fundamental tensions in trademark law — the conflict between registered rights (which carry statutory presumption of validity) and unregistered prior use rights (which are based on common law principles and the equitable doctrine that registration cannot extinguish pre-existing legitimate use). X's legal position is governed by Sections 34, 35, and related provisions of the Trade Marks Act, 1999, and must be analyzed across two fronts: (a) defense to Y's infringement suit, and (b) grounds to challenge Y's registration."),
  subHeading("2. X's Defense of Prior Use — Section 34"),
  body("Section 34 of the Trade Marks Act, 1999 is the cornerstone of X's defense. It provides that where a person other than the registered proprietor of a trademark has been using the mark continuously from before the date on which the registered proprietor or their predecessor in title first used the mark, the registered proprietor cannot interfere with or restrain the use of the mark by that prior user."),
  body("For X to successfully invoke Section 34: (a) X must establish that they have used the mark continuously for 15 years — predating Y's first use; (b) X must prove use in relation to the same or similar goods or services; and (c) X must demonstrate that Y's use commenced after X's. If X can establish these facts, X has a complete statutory defense against Y's infringement suit, and Y's registration does not override X's prior use rights."),
  body("The crucial date under Section 34 is the date from which the registered proprietor or their predecessor first used the mark, not the date of registration. If Y registered the mark but did not personally use it before X's use, Section 34 protects X even though Y holds the registration certificate."),
  subHeading("3. Prior Use in the Context of the Registration Process — Section 12"),
  body("Section 12 of the Act provides that a trademark may be registered despite an identical or similar existing mark if the Registrar is satisfied that there has been honest concurrent use. X's prior use is directly relevant if there was a procedure for X to be recorded as an honest concurrent user. If Y obtained registration without the Registrar being aware of X's prior use, the registration may be challengeable."),
  subHeading("4. Challenging Y's Registration — Rectification Proceedings"),
  body("In addition to raising prior use as a defense in the infringement suit, X has strong grounds to challenge Y's registration before the Intellectual Property Appellate Board (IPAB) or the High Court under Section 57 of the Trade Marks Act, 1999, which provides for rectification of the Register."),
  subHeading("Grounds for Rectification:"),
  bullet("Section 11(1) — Relative Grounds: At the time Y applied for registration, X's mark was being used and had acquired a reputation and goodwill in the market for 15 years. Y's registration of an identical mark for the same goods is a relative ground for refusal. X should have been identified in the trademark search, and failure to identify X's prior use may be ground for rectification."),
  bullet("Section 47 — Non-Use by Y: If Y obtained the registration but is not genuinely using the trademark in India in connection with goods or services (perhaps Y registered the mark speculatively after seeing X's use), X can apply for cancellation of Y's registration on the ground of non-use for a continuous period of five years after registration."),
  bullet("Section 9/11 — Registration Obtained in Bad Faith: If Y registered the mark with knowledge of X's long-standing prior use (which is plausible given 15 years of X's use), the registration was obtained in bad faith. Under Section 11(10)(ii), bad faith at the time of application is a ground for refusal of registration and, post-registration, for rectification. Indian courts have increasingly recognised bad faith as a ground for challenging trademark registrations."),
  subHeading("5. Well-Known Mark Argument"),
  body("If X's 15-year prior use has been extensive, geographically widespread, and has resulted in the mark achieving significant recognition among consumers, X may apply to the Registrar for recognition of the mark as a 'well-known trademark' under Section 11(6). Recognition as a well-known mark provides expanded cross-category protection and can be used as a weapon in the rectification proceedings against Y's registration."),
  subHeading("6. Passing Off"),
  body("Even if X's defense under Section 34 is limited in scope (for example, if Y's registration predates X's use in certain territories), X can independently pursue a passing off action against Y if Y's use of the registered mark misleads the public into thinking Y's goods are X's goods or vice versa. X's 15-year use has created goodwill and reputation, and Y's registered trademark — if used in a manner that causes confusion with X's market presence — constitutes misrepresentation causing damage to X's goodwill."),
  subHeading("7. Practical Strategy for X"),
  body("X should take the following steps: Immediately file a rectification application before the IPAB (or High Court) under Section 57, challenging Y's registration on grounds of prior use, bad faith, and relative grounds; File a counter-claim in Y's infringement suit asserting the Section 34 prior use defense; Gather and preserve evidence of continuous use — invoices, advertising materials, customer testimonials, press coverage, and any market surveys dating back 15 years; File a trademark application for the mark to establish X's own registered rights going forward, which will be protected from the date of application; Explore whether Y's registration can be attacked on bad faith grounds."),
  subHeading("8. Conclusion"),
  body("X's legal position under the Trade Marks Act, 1999 is reasonably strong. Section 34 provides a direct statutory defense against Y's infringement suit, provided X can prove continuous prior use predating Y's use. X also has strong grounds to challenge Y's registration through rectification proceedings under Section 57, particularly on the grounds of Y's bad faith registration with knowledge of X's prior use and Y's possible non-use of the mark. The 15-year prior use, if well-documented, is X's most powerful weapon both as a defense and as an offensive tool to cancel Y's registration."),
  spacer()
);

// B-Q15
children.push(
  qHeading(15, 10, "'Unfair competition as a concept is broader than passing off and trademark infringement put together.' Critically examine this statement comparing scope, elements, and remedies."),
  ansLabel(),
  subHeading("1. Introduction"),
  body("Unfair competition is a broad, elastic concept encompassing any commercial practice that is dishonest, deceptive, or contrary to good faith commercial standards, causing harm to competitors or consumers. The statement that 'unfair competition is broader than passing off and trademark infringement put together' merits critical scrutiny. While there is merit in the assertion — particularly in jurisdictions that have developed a comprehensive unfair competition framework — the position in India is more nuanced, as Indian law does not have a standalone unfair competition statute but addresses many of its manifestations through various statutory and common law mechanisms."),
  subHeading("2. Trademark Infringement — Scope and Elements"),
  body("Trademark infringement, as discussed, is a statutory wrong under the Trade Marks Act, 1999, requiring: a registered trademark; unauthorized use of an identical or similar mark; in the course of trade; for identical or similar goods/services (or, for well-known marks, even dissimilar goods); creating a likelihood of confusion. Its scope is narrow in the sense that it is limited to registered marks. A trader with an unregistered mark, or one whose mark has been appropriated in a way that does not technically qualify as 'use in the course of trade', may find trademark infringement law inadequate."),
  subHeading("3. Passing Off — Scope and Elements"),
  body("Passing off, as established, is broader than trademark infringement because it protects unregistered marks and gets-up, trade names, and any business identifier in which goodwill subsists. Its three-element classical test (goodwill, misrepresentation, damage) is flexible. Extended passing off further protects collective reputations in product categories. However, passing off still has limitations: it requires a competitive relationship (or at least a potential for consumer confusion); it requires actual goodwill in the jurisdiction; and it does not cover wrongs that do not involve a misrepresentation (such as direct business sabotage, bribery of competitors' employees, or trade secret misappropriation)."),
  subHeading("4. Unfair Competition — Broader Concept"),
  body("'Unfair competition' as a concept encompasses a wider range of commercial wrongs, including:"),
  bullet("Misappropriation of trade secrets and confidential business information."),
  bullet("Comparative advertising that is false, misleading, or disparaging."),
  bullet("Bribery of competitors' employees to obtain trade secrets."),
  bullet("Interference with contractual relations or prospective economic advantage."),
  bullet("Misuse of a competitor's confidential customer list."),
  bullet("Slavish imitation of product design or get-up that is not trademark-protected."),
  bullet("Making false statements about a competitor's goods or services (trade libel/malicious falsehood)."),
  bullet("Domain name cybersquatting and other forms of online misappropriation."),
  bullet("Palming off, reverse palming off, and initial interest confusion."),
  body("None of these wrongs necessarily involves a registered trademark. Some (like slavish imitation without trademark significance) may not satisfy the elements of passing off either, yet they are commercially damaging and ethically objectionable. To this extent, unfair competition is indeed broader than the sum of passing off and trademark infringement."),
  subHeading("5. The Indian Position"),
  body("India does not have a standalone Unfair Competition Act. The concept is addressed through: the Trade Marks Act, 1999 (infringement and dilution); the common law tort of passing off; the Copyright Act, 1957 (misappropriation of creative expression); the Trade Secrets Protection (through contractual remedies and the Specific Relief Act); the Consumer Protection Act, 2019 (protecting consumers against misleading commercial practices); the Competition Act, 2002 (prohibiting anti-competitive practices); and the Indian Penal Code (for fraud and cheating)."),
  body("Article 10bis of the Paris Convention obliges member states (including India) to provide effective protection against unfair competition, defined as any act of competition contrary to honest practices in industrial or commercial matters. Despite this international obligation, India has not enacted a comprehensive Unfair Competition statute, relying instead on this patchwork of laws."),
  subHeading("6. Critical Analysis of the Statement"),
  body("The statement is substantially correct in a comparative law sense. In jurisdictions like Germany (which has the comprehensive Gesetz gegen unlauteren Wettbewerb — UWG) and the United States (which has broad Lanham Act protections under Section 43(a) supplemented by state unfair competition law), unfair competition law provides remedies for a far wider range of commercial wrongs than passing off or trademark infringement alone can cover."),
  body("In India, however, the statement is aspirationally rather than operationally true. The lack of a standalone unfair competition statute means that Indian courts must stretch passing off, copyright law, consumer protection law, and contract law to cover wrongs that a unified unfair competition statute would address more directly. This creates uncertainty, inconsistency, and inadequate remedies for novel forms of commercial wrongdoing — particularly in the digital economy."),
  body("That said, the statement must be qualified: passing off in its extended form is itself quite broad; Indian courts have shown creativity in applying existing laws to new commercial wrongs; and the Competition Act, 2002 addresses market-distorting practices that go beyond the scope of any single IP law."),
  subHeading("7. Suggested Legislative Response"),
  body("India would benefit from a comprehensive Unfair Competition Act that: defines unfair competition to cover all dishonest commercial practices contrary to good faith; explicitly includes misappropriation of trade secrets, false advertising, slavish imitation, and online misappropriation; provides both civil remedies (injunction, damages, account of profits) and administrative remedies; and aligns with India's obligations under Article 10bis of the Paris Convention."),
  subHeading("8. Conclusion"),
  body("The statement that 'unfair competition is broader than passing off and trademark infringement put together' is analytically sound. Unfair competition captures a wider universe of commercial wrongdoing — from trade secret theft to false advertising to slavish imitation — that neither trademark infringement (which requires registration) nor passing off (which requires misrepresentation and goodwill) can fully address. In India, the gap left by the absence of a standalone unfair competition statute is partially bridged by a combination of laws, but a comprehensive legislative reform remains necessary to provide adequate, coherent, and modern protection against unfair competition in all its forms."),
  spacer()
);

// ─── BUILD ─────────────────────────────────────────────────────────────────

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "1F3864" },
        paragraph: { spacing: { before: 360, after: 240 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "1F3864" },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/LAW352_Answer_Key.docx', buf);
  console.log('Done');
}).catch(err => { console.error(err); process.exit(1); });