const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel,
  PageBreak, LevelFormat, VerticalAlign
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 8, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

const CONTENT_W = 9360;
const COL1 = 2800;
const COL2 = CONTENT_W - COL1;

function heading(text, size = 32, center = true) {
  return new Paragraph({
    alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: 160, after: 80 },
    children: [new TextRun({ text, bold: true, size, font: "Arial" })]
  });
}

function subheading(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 60 },
    children: [new TextRun({ text, bold: true, size: 22, font: "Arial" })]
  });
}

function body(text, center = false, italic = false) {
  return new Paragraph({
    alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 20, font: "Arial", italics: italic })]
  });
}

function bold(text, center = false) {
  return new Paragraph({
    alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: 60, after: 40 },
    children: [new TextRun({ text, bold: true, size: 20, font: "Arial" })]
  });
}

function cell(children, width, shade = null, valign = VerticalAlign.TOP) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: shade ? { fill: shade, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: valign,
    children
  });
}

function answerCell(text, width) {
  return cell([
    new Paragraph({
      spacing: { before: 40, after: 40 },
      children: [new TextRun({ text, size: 20, font: "Arial", italics: true, color: "1F4E79" })]
    })
  ], width);
}

function emptyCell(width, lines = 3) {
  const children = [];
  for (let i = 0; i < lines; i++) {
    children.push(new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: "", size: 20 })] }));
  }
  return cell(children, width);
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function taskRow(questionText, answerText, lines = 3) {
  return new TableRow({
    children: [
      cell([
        new Paragraph({
          spacing: { before: 40, after: 40 },
          children: [new TextRun({ text: questionText, size: 20, font: "Arial" })]
        })
      ], COL1),
      answerText ? answerCell(answerText, COL2) : emptyCell(COL2, lines)
    ]
  });
}

// ─── PAGE 1: TED Talk analysis sheet ───────────────────────────────────────
const page1 = [
  heading("ANALYZING TED TALKS FOR:", 28),
  heading("ethos, pathos, and logos", 36),
  new Paragraph({ spacing: { before: 40, after: 40 }, children: [] }),
  subheading("OBJECTIVE"),
  body("To recognize, differentiate, and analyze rhetorical devices (Ethos, Pathos, Logos) used in speeches and spoken language.", true),
  new Paragraph({ spacing: { before: 80, after: 40 }, children: [] }),
  subheading("YOUR TASK"),
  body("While watching a TED talk presentation, complete the table below as the information and rhetorical devices are presented.", true),
  new Paragraph({ spacing: { before: 80, after: 40 }, children: [] }),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [COL1, COL2],
    rows: [
      taskRow("Evaluate the opening of the speech. What effective techniques does the speaker use?", "", 4),
      taskRow("Evaluate the use of ethical appeal. How does the speaker establish ETHOS?", "", 4),
      taskRow("Evaluate the use of emotional appeal. How does the speaker establish PATHOS?", "", 4),
      taskRow("Evaluate the use of logical appeal. How does the speaker establish LOGOS?", "", 4),
      taskRow("Evaluate the closing of the speech. What effective techniques does the speaker use?", "", 4),
      taskRow("Analyze the presentation of the speech (not the content), and evaluate what the speaker did to present an overall, effective speech", "", 5),
    ]
  }),
  pageBreak()
];

// ─── PAGE 2: ETHOS ─────────────────────────────────────────────────────────

const THIRD = Math.floor(CONTENT_W / 3);
const TWOTHIRD = CONTENT_W - THIRD;

function infoBox(leftLabel, leftText, middleContent, rightLabel, rightItems) {
  // 3 col layout using a table with no borders
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [2200, 4960, 2200],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: border, bottom: border, left: border, right: border },
            width: { size: 2200, type: WidthType.DXA },
            shading: { fill: "D9D9D9", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [
              new Paragraph({ children: [new TextRun({ text: leftLabel, bold: true, size: 20, font: "Arial" })] }),
              ...leftText.map(t => new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: t, size: 18, font: "Arial" })] }))
            ]
          }),
          new TableCell({
            borders: { top: border, bottom: border, left: border, right: border },
            width: { size: 4960, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: middleContent
          }),
          new TableCell({
            borders: { top: border, bottom: border, left: border, right: border },
            width: { size: 2200, type: WidthType.DXA },
            shading: { fill: "D9D9D9", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [
              new Paragraph({ children: [new TextRun({ text: rightLabel, bold: true, size: 20, font: "Arial" })] }),
              ...rightItems.map(t => new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: t, size: 18, font: "Arial" })] }))
            ]
          })
        ]
      })
    ]
  });
}

const ethosMiddle = [
  new Paragraph({ children: [new TextRun({ text: "LET'S FOCUS ON ETHOS", bold: true, size: 28, font: "Arial" })] }),
  new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: "How can you incorporate ethos?", bold: true, size: 20, font: "Arial", italics: true })] }),
  new Paragraph({ spacing: { before: 40, after: 20 }, children: [new TextRun({ text: "• You can openly remind your audience who you are and why you are an authority on the subject. ('As the leading researcher, I agree...')", size: 18, font: "Arial" })] }),
  new Paragraph({ spacing: { before: 20, after: 20 }, children: [new TextRun({ text: "• You can establish authority more subtly through the use of jargon or specialized terms ('My quantitative research in linguistic anthropology has shown...').", size: 18, font: "Arial" })] }),
  new Paragraph({ spacing: { before: 20, after: 20 }, children: [new TextRun({ text: "• You can establish credibility by just using correct grammar and language, using solid reasoning and good arguments and therefore sounding credible and trustworthy.", size: 18, font: "Arial" })] }),
];

const ethosRightItems = [
  "EXAMPLES:",
  "» I have been married for 58 years and I can tell you that he will not be a good husband to you.",
  "» Having written ten successful novels myself, I can tell you that this book is worth buying.",
  "» My friend, who has a PhD in nutritional science, says that we should eat fewer carbohydrates."
];

const whyEthos = new Paragraph({
  spacing: { before: 80, after: 60 },
  children: [
    new TextRun({ text: "Why use ethos?  ", bold: true, size: 20, font: "Arial", italics: true }),
    new TextRun({ text: "If your audience believe you to be a qualified, authoritative figure, or an expert in a subject, they are more likely to be persuaded by your argument.", size: 20, font: "Arial" })
  ]
});

// Task One table for ethos
const E_Q = 1800, E_Y = 900, E_EX = CONTENT_W - E_Q - E_Y;

const ethosTask1Rows = [
  {
    q: '"You should definitely get help with your debt, & I say that as somebody with over 30 years experience of debt counseling."',
    ethos: "Yes",
    explain: "The speaker has referred to their extensive experience in this area, in order to give credit to their opinion. Somebody with over 30 years experience would surely know what he or she is talking about.",
    filled: true
  },
  {
    num: "1",
    q: '"As your father, I love you and only want the best for you. Therefore when I ask you not to go, please listen to me."',
    ethos: "Yes",
    explain: "you are an authority to the subject",
    filled: true, handwritten: true
  },
  {
    num: "2",
    q: '"We really should try that recipe. Someone told me that it was good."',
    ethos: "No",
    explain: "personal opinion doesn't show we solid arguments",
    filled: true, handwritten: true
  },
  {
    num: "3",
    q: '"The research – conducted by professors at Harvard University – suggests that you should learn a second language."',
    ethos: "Yes",
    explain: "The study mentioned was done by educators at a prestigious university",
    filled: true, handwritten: true
  },
  {
    num: "4",
    q: '"Dentists all over the world are telling their patients the same things. You must floss regularly."',
    ethos: "Yes",
    explain: "Trusted sources individuals giving advice about a topic they specialise in.",
    filled: true, handwritten: true
  },
  {
    num: "5",
    q: '"I read somewhere that bicarbonate of soda is really useful for cleaning. You should try it."',
    ethos: "No",
    explain: "Unclear source of information that may affect infrastructure",
    filled: true, handwritten: true
  },
];

const ethosTask1 = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [E_Q, E_Y, E_EX],
  rows: [
    new TableRow({
      children: [
        cell([new Paragraph({ children: [new TextRun({ text: "", size: 20 })] })], E_Q, "BFBFBF"),
        cell([new Paragraph({ children: [new TextRun({ text: "Ethos?", bold: true, size: 20, font: "Arial" })] })], E_Y, "BFBFBF"),
        cell([new Paragraph({ children: [new TextRun({ text: "Explain", bold: true, size: 20, font: "Arial" })] })], E_EX, "BFBFBF"),
      ]
    }),
    ...ethosTask1Rows.map(r => new TableRow({
      children: [
        cell([
          new Paragraph({
            spacing: { before: 30, after: 30 },
            children: [
              r.num ? new TextRun({ text: r.num + "  ", bold: true, size: 20, font: "Arial" }) : new TextRun({ text: "" }),
              new TextRun({ text: r.q, size: 20, font: "Arial", italics: true })
            ]
          })
        ], E_Q),
        cell([new Paragraph({ children: [new TextRun({ text: r.ethos, size: 20, font: "Arial", color: r.handwritten ? "1F4E79" : "000000", italics: r.handwritten })] })], E_Y),
        cell([new Paragraph({ children: [new TextRun({ text: r.explain, size: 20, font: "Arial", color: r.handwritten ? "1F4E79" : "000000", italics: r.handwritten })] })], E_EX),
      ]
    }))
  ]
});

function writingLines(n = 3) {
  const lines = [];
  for (let i = 0; i < n; i++) {
    lines.push(new Paragraph({
      spacing: { before: 60, after: 0 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" } },
      children: [new TextRun({ text: "", size: 20 })]
    }));
  }
  return lines;
}

const page2 = [
  heading("PERSUASIVE LANGUAGE", 20),
  infoBox(
    "Ethos is an appeal to ethics; it is a means of convincing the audience of the character or credibility of the persuader.",
    [],
    ethosMiddle,
    "EXAMPLES:",
    ethosRightItems.slice(1)
  ),
  whyEthos,
  bold("TASK ONE: For each of the following, identify whether or not ETHOS is used, and explain your reasoning."),
  ethosTask1,
  new Paragraph({ spacing: { before: 120, after: 40 }, children: [] }),
  bold("TASK TWO: For each of the following scenarios, write your own persuasive sentences using ethos:"),
  body("1. You are trying to persuade your mother to let you get a tattoo."),
  ...writingLines(3),
  new Paragraph({ spacing: { before: 60, after: 40 }, children: [] }),
  body("2. You are attempting to convince an audience that animal testing is morally wrong."),
  ...writingLines(3),
  new Paragraph({ spacing: { before: 60, after: 40 }, children: [] }),
  body("3. You are trying to persuade your teacher not to give you homework."),
  ...writingLines(3),
  pageBreak()
];

// ─── PAGE 3: PATHOS ────────────────────────────────────────────────────────

const pathosMiddle = [
  new Paragraph({ children: [new TextRun({ text: "LET'S FOCUS ON PATHOS", bold: true, size: 28, font: "Arial" })] }),
  new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: "How can you incorporate pathos?", bold: true, size: 20, font: "Arial", italics: true })] }),
  new Paragraph({ spacing: { before: 40, after: 20 }, children: [new TextRun({ text: "• You can use descriptive language and imagery, which evokes emotions. ('This life-saving dishwasher will enhance your life and lighten your load.')", size: 18, font: "Arial" })] }),
  new Paragraph({ spacing: { before: 20, after: 20 }, children: [new TextRun({ text: "• You can identify values / emotions which relate particularly to your audience, and play on them. ('I know you care for your family, so...')", size: 18, font: "Arial" })] }),
  new Paragraph({ spacing: { before: 20, after: 20 }, children: [new TextRun({ text: "• You can use personal stories to appeal to the sympathies and emotions of your audience. ('When I lost my job, I knew how important it was to...')", size: 18, font: "Arial" })] }),
];

const pathosRightItems = [
  "» If you don't go on this holiday you will regret it. You don't want to live with regrets, do you?",
  "» You love your dog; so buy this dog food, as it will help him to get all his nutrients.",
  "» You have been poorly treated. You have been taken advantage of. Now is the time to take action."
];

const whyPathos = new Paragraph({
  spacing: { before: 80, after: 60 },
  children: [
    new TextRun({ text: "Why use pathos?  ", bold: true, size: 20, font: "Arial", italics: true }),
    new TextRun({ text: "If your audience is emotionally invested and engaged with your cause or argument, they are more likely to be persuaded.", size: 20, font: "Arial" })
  ]
});

const P_Q = 2500, P_EX = CONTENT_W - P_Q;

const pathosTask1Rows = [
  {
    q: '"Don\'t be the last person to get one. You don\'t want to be the laughing stock of your school!"',
    explain: "Here the speaker is appealing to the audience's sense of pride by pointing out that they will be laughed at if they don't get the product. This would make the listener/reader feel compelled to get it, in order to maintain their dignity and not hurt their pride.",
    filled: true
  },
  { num: "1", q: '"If you don\'t purchase this life insurance, and something happens to you, how will your family survive?"', explain: "", filled: false },
  { num: "2", q: '"We have been mistreated, abused and oppressed. They have benefited from our suffering and we must act now!"', explain: "", filled: false },
  { num: "3", q: '"We live in a great democracy. So donate now and support the troops who are protecting our freedom."', explain: "", filled: false },
  { num: "4", q: '"Just eat and don\'t complain. Children in Africa are starving and would give anything to have that plate of food."', explain: "", filled: false },
  { num: "5", q: '"Caring for the environment may not change your life, but it will change the lives of your children."', explain: "", filled: false },
];

const pathosTask1 = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [P_Q, P_EX],
  rows: [
    new TableRow({
      children: [
        cell([new Paragraph({ children: [new TextRun({ text: "", size: 20 })] })], P_Q, "BFBFBF"),
        cell([new Paragraph({ children: [new TextRun({ text: "Explain", bold: true, size: 20, font: "Arial" })] })], P_EX, "BFBFBF"),
      ]
    }),
    ...pathosTask1Rows.map(r => new TableRow({
      children: [
        cell([
          new Paragraph({
            spacing: { before: 30, after: 30 },
            children: [
              r.num ? new TextRun({ text: r.num + "  ", bold: true, size: 20, font: "Arial" }) : new TextRun({ text: "" }),
              new TextRun({ text: r.q, size: 20, font: "Arial", italics: true })
            ]
          })
        ], P_Q),
        r.filled
          ? cell([new Paragraph({ spacing: { before: 40 }, children: [new TextRun({ text: r.explain, size: 20, font: "Arial" })] })], P_EX)
          : emptyCell(P_EX, 3)
      ]
    }))
  ]
});

const page3 = [
  heading("PERSUASIVE LANGUAGE", 20),
  infoBox(
    "Pathos is an appeal to emotion; it is a means of convincing the audience by creating an emotional response.",
    [],
    pathosMiddle,
    "EXAMPLES:",
    pathosRightItems
  ),
  whyPathos,
  bold("TASK ONE: For each of the following, explain how pathos is used (you must identify the emotional appeal)."),
  pathosTask1,
  new Paragraph({ spacing: { before: 120, after: 40 }, children: [] }),
  bold("TASK TWO: For each of the following scenarios, write your own persuasive sentences using pathos:"),
  body("1. You are trying to persuade members of your class to volunteer to tutor underprivileged children."),
  ...writingLines(3),
  new Paragraph({ spacing: { before: 60, after: 40 }, children: [] }),
  body("2. You are attempting to convince a friend to join your sports team."),
  ...writingLines(3),
  new Paragraph({ spacing: { before: 60, after: 40 }, children: [] }),
  body("3. You are trying to persuade an audience that they should buy a particular brand of cereal."),
  ...writingLines(3),
  pageBreak()
];

// ─── PAGE 4: LOGOS ─────────────────────────────────────────────────────────

const logosMiddle = [
  new Paragraph({ children: [new TextRun({ text: "LET'S FOCUS ON LOGOS", bold: true, size: 28, font: "Arial" })] }),
  new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: "How can you incorporate logos?", bold: true, size: 20, font: "Arial", italics: true })] }),
  new Paragraph({ spacing: { before: 40, after: 20 }, children: [new TextRun({ text: "• You can use research and statistics to back-up your arguments. ('87% of students benefit from textbooks, so...')", size: 18, font: "Arial" })] }),
  new Paragraph({ spacing: { before: 20, after: 20 }, children: [new TextRun({ text: "• You can use logical processes to explain your point – 'if this... then that...' ('If you don't want gum disease, then you should brush regularly.')", size: 18, font: "Arial" })] }),
  new Paragraph({ spacing: { before: 20, after: 20 }, children: [new TextRun({ text: "• You must give convincing, solid evidence and reasons to support your claims. ('I need a car; it would make me more independent; save me money; improve my life.')", size: 18, font: "Arial" })] }),
];

const logosRightItems = [
  "» I've not eaten meat for 8 years, and I'm fit and healthy; you can't argue that vegetarianism is always unhealthy.",
  "» We have conducted the experiment 57 times and we get the same results every single time.",
  "» That is not my wallet. My wallet has a tear in it. This wallet has no tear in it. Therefore, it cannot be mine."
];

const whyLogos = new Paragraph({
  spacing: { before: 80, after: 60 },
  children: [
    new TextRun({ text: "Why use logos?  ", bold: true, size: 20, font: "Arial", italics: true }),
    new TextRun({ text: "If you can present a logical, rational argument to your audience, which engages their intellect and sense of reason, they are more likely to be persuaded.", size: 20, font: "Arial" })
  ]
});

const L_Q = 2500, L_Y = 600, L_N = 600, L_EX = CONTENT_W - L_Q - L_Y - L_N;

const logosTask1Rows = [
  {
    q: '"I need new jeans. Everyone in my class has new jeans."',
    effective: false,
    explain: "This isn't an effective use of logos as the reasoning is not logical and convincing. Firstly, it's highly unlikely that everyone has new jeans (hyperbole). Secondly, just because other people have new jeans, it doesn't logically follow that the speaker needs them.",
    filled: true
  },
  { num: "1", q: '"All men and women will die. You are a man. Therefore, you will die one day."', effective: null, explain: "", filled: false },
  { num: "2", q: '"Everyone has children. Therefore, everyone needs to think about the schooling of his or her children."', effective: null, explain: "", filled: false },
  { num: "3", q: '"You don\'t need to jump in front of a train to know it\'s a bad idea; so why do you need to try drugs to know if they\'re damaging?"', effective: null, explain: "", filled: false },
  { num: "4", q: '"Cigarette smoke contains over 4,800 chemicals, 69 of which are known to cause cancer. So why start smoking?"', effective: null, explain: "", filled: false },
  { num: "5", q: '"Every morning the rooster crows, then the sun rises. Therefore the rooster causes the sun to rise."', effective: null, explain: "", filled: false },
];

const logosTask1 = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [L_Q, L_Y, L_N, L_EX],
  rows: [
    new TableRow({
      children: [
        cell([new Paragraph({ children: [new TextRun({ text: "", size: 20 })] })], L_Q, "BFBFBF"),
        new TableCell({
          borders,
          width: { size: L_Y + L_N, type: WidthType.DXA },
          shading: { fill: "BFBFBF", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          columnSpan: 2,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Effective?", bold: true, size: 20, font: "Arial" })] })]
        }),
        cell([new Paragraph({ children: [new TextRun({ text: "Explain", bold: true, size: 20, font: "Arial" })] })], L_EX, "BFBFBF"),
      ]
    }),
    new TableRow({
      children: [
        cell([new Paragraph({ children: [new TextRun({ text: "", size: 20 })] })], L_Q),
        cell([new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Yes", bold: true, size: 20, font: "Arial" })] })], L_Y),
        cell([new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "No", bold: true, size: 20, font: "Arial" })] })], L_N),
        cell([new Paragraph({ children: [new TextRun({ text: "", size: 20 })] })], L_EX),
      ]
    }),
    ...logosTask1Rows.map(r => new TableRow({
      children: [
        cell([
          new Paragraph({
            spacing: { before: 30, after: 30 },
            children: [
              r.num ? new TextRun({ text: r.num + "  ", bold: true, size: 20, font: "Arial" }) : new TextRun({ text: "" }),
              new TextRun({ text: r.q, size: 20, font: "Arial", italics: true })
            ]
          })
        ], L_Q),
        cell([new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: r.filled && r.effective === true ? "✓" : "", size: 20, font: "Arial" })] })], L_Y),
        cell([new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: r.filled && r.effective === false ? "✓" : "", size: 20, font: "Arial" })] })], L_N),
        r.filled
          ? cell([new Paragraph({ spacing: { before: 40 }, children: [new TextRun({ text: r.explain, size: 20, font: "Arial" })] })], L_EX)
          : emptyCell(L_EX, 3)
      ]
    }))
  ]
});

const page4 = [
  heading("PERSUASIVE LANGUAGE", 20),
  infoBox(
    "Logos is an appeal to logic; it is a means of convincing the audience through rational thought and reason.",
    [],
    logosMiddle,
    "EXAMPLES:",
    logosRightItems
  ),
  whyLogos,
  bold("TASK ONE: For each of the following, explain whether or not you think that logos is being used effectively."),
  logosTask1,
  new Paragraph({ spacing: { before: 120, after: 40 }, children: [] }),
  bold("TASK TWO: For each of the following scenarios, write your own persuasive sentences using logos:"),
  body("1. You are trying to persuade your brother to stop eating chocolate."),
  ...writingLines(3),
  new Paragraph({ spacing: { before: 60, after: 40 }, children: [] }),
  body("2. You are attempting to convince your class that they shouldn't drink and drive."),
  ...writingLines(3),
  new Paragraph({ spacing: { before: 60, after: 40 }, children: [] }),
  body("3. You are trying to persuade an audience that they should exercise more often."),
  ...writingLines(3),
  pageBreak()
];

// ─── PAGE 5: USING ETHOS, PATHOS & LOGOS ───────────────────────────────────

function scenario(num, text, answerText = "") {
  const paras = [
    new Paragraph({
      spacing: { before: 80, after: 40 },
      children: [new TextRun({ text: `${num}. ${text}`, size: 20, font: "Arial" })]
    }),
    ...writingLines(2),
    new Paragraph({
      spacing: { before: 60, after: 20 },
      children: [new TextRun({ text: "Technique(s) used: ", size: 20, font: "Arial", italics: true })]
    })
  ];
  return paras;
}

const page5 = [
  heading("PERSUASIVE LANGUAGE", 20),
  heading("USING ETHOS, PATHOS & LOGOS", 28),
  new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text: "For each of the following situations, write your own persuasive sentence(s) using ethos, pathos or logos. You may use more than one technique per sentence, but you must use all of them at least once at some point in the exercise. You should also identify which one(s) you have tried to implement. The first one has been done for you as an example.", size: 20, font: "Arial" })]
  }),
  new Paragraph({
    spacing: { before: 40, after: 20 },
    children: [new TextRun({ text: "Example. You are trying to sell a particular brand of toothpaste, called 'ShinyWhite'.", size: 20, font: "Arial", bold: true })]
  }),
  new Paragraph({
    spacing: { before: 20, after: 20 },
    children: [new TextRun({ text: "3 out of 4 people surveyed said that they prefer our brand to any other. If you care about your teeth, make the wise choice and buy ShinyWhite before it's too late.", size: 20, font: "Arial", italics: true, color: "1F4E79" })]
  }),
  new Paragraph({
    spacing: { before: 20, after: 40 },
    children: [
      new TextRun({ text: "Technique(s) used: ", size: 20, font: "Arial", italics: true }),
      new TextRun({ text: "Pathos and Logos", size: 20, font: "Arial", italics: true, color: "1F4E79" })
    ]
  }),
  ...scenario("1", "The principal has cancelled the prom. You are trying to persuade him/her to reconsider his/her decision."),
  ...scenario("2", "You are attempting to convince your parents to buy you a pet rabbit."),
  ...scenario("3", "You are trying to persuade your aunt to set up a Facebook account."),
  ...scenario("4", "Your sister is on a strict diet, but you are trying to persuade her to try one of your freshly baked cookies."),
  ...scenario("5", "It is a rainy day; you are trying to persuade your friend to go for a run with you."),
  ...scenario("6", "Your mother wants you to go grocery shopping with her and you are trying to persuade her that it's not a good idea."),
  ...scenario("7", "Your teacher wants you to handwrite your essay but you want to type it; try to persuade him/her to let you type it."),
  ...scenario("8", "You are trying to convince your classmates to each donate $5 to a charity which rescues abandoned animals."),
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 720, right: 720, bottom: 720, left: 720 }
      }
    },
    children: [
      ...page1,
      ...page2,
      ...page3,
      ...page4,
      ...page5
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/ethos_pathos_logos_worksheet.docx', buf);
  console.log('Done!');
});
