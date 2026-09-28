// STEAM教育 競合比較スライド生成
const pptxgen = require("pptxgenjs");
const path = require("path");
const D = require("./data.json");

const OUT = path.join(__dirname, "..", "STEAM教育_競合比較_サマリー.pptx");
const FONT = "Meiryo";
const C = {
  navy: "1F2A44", yel: "F5B700", ink: "222222", sub: "5B6475", line: "D5DAE3",
  pale: "F3F5F9", paleY: "FFF6D6", white: "FFFFFF",
  robo: "2F6DB5", prog: "7A4FB5", online: "2E9C8F",
  hi: "C62828", mid: "D98200", lo: "2E7D32",
};
const CAT_COLOR = { "自社": C.yel, "ロボット・ものづくり": C.robo, "プログラミング特化": C.prog, "オンライン・通信": C.online };
const THREAT_COLOR = { "高": C.hi, "中": C.mid, "低": C.lo, "—": C.sub };
const COS = D.companies;
const byKey = Object.fromEntries(COS.map(c => [c.key, c]));
const man = v => (v / 10000).toFixed(1);

// ---- スライド用の短縮テキスト ----
const S = {
  stemon:      { lesson: "60分×月3〜4回（幼児 月2回）", size: "幼児4名・小学生6名", cur: "しる→つくる→ためす。物理の仕組み＋プログラミング＋ロボット", contest: "ステモンカップ（小学生ロボコン予選）", fb: "セミナーあり／毎回レポートは要確認",
                 emp: "アルバイト", wage: "コマ給2,000〜2,700円\n（時給換算 約1,330〜1,800円）", train: "2か月（座学＋実地）", ratio: "1:4〜6", trev: "楽しく通える／個別対応が少ない",
                 mat: "独自カリキュラム＋ブロック・PC", mtype: "独自", mcost: "貸出（購入不要）", good: "夢中になる・仕組みから学べる・ロボコン", bad: "高め・個別対応が少ない・地方に少ない" },
  steamcampus: { lesson: "50〜90分（回数は要確認）", size: "4〜8名（教室差）", cur: "LEGOで年齢別の1年コース（2歳〜）", contest: "特になし", fb: "要確認",
                 emp: "アルバイト", wage: "1,200〜1,600円", train: "要確認", ratio: "4〜8名", trev: "質が高い／運営会社で差",
                 mat: "LEGO Education", mtype: "既製", mcost: "毎年購入（1〜4.7万円）", good: "質の高いプログラム・幼児から", bad: "教材を毎年購入・体験有料・ブランド混乱",
                 issue: "LEGOブランド喪失、教材の毎年購入", attack: "ブランドの一貫性、購入不要で総額が明快" },
  human:       { lesson: "90分×月2回", size: "最大10名程度", cur: "ロボット製作→改造・プログラミング（最長8年）", contest: "ロボット教室全国大会", fb: "授業最後の発表",
                 emp: "業務委託", wage: "1,320円〜", train: "本部マニュアル", ratio: "〜10名", trev: "教室によってばらつき",
                 mat: "独自ロボットキット", mtype: "独自", mcost: "購入 33,000円＋進級時追加", good: "近所にある・ロボットがかっこいい", bad: "講師差・組立中心・初期費用",
                 issue: "講師品質のばらつき、組立中心、月2回", attack: "講師研修、探究型、授業回数の多さ" },
  robodone:    { lesson: "50〜120分×月3回程度", size: "講師1〜3名体制", cur: "SPIKEでペア学習＋毎回の発表", contest: "社内大会・WRO実績", fb: "アプリで学習状況を共有",
                 emp: "アルバイト", wage: "要確認", train: "要確認", ratio: "要確認", trev: "明るい／理解度が見えにくい",
                 mat: "LEGO SPIKE Prime", mtype: "既製", mcost: "貸与（持ち帰り不可）", good: "ペア学習と発表・アプリ", bad: "持ち帰り不可・隠れ費用",
                 issue: "成果が家に残らない、料金体系が複雑", attack: "作品持ち帰り、保護者レポート" },
  crefus:      { lesson: "90分×年42回", size: "要確認", cur: "学年別テーマ、理科・算数と連動", contest: "FLL世界大会に出場（2026年）", fb: "要確認",
                 emp: "アルバイト", wage: "2,000円以上", train: "研修あり・力量で担当", ratio: "要確認", trev: "ハキハキ／専門性に差",
                 mat: "LEGOベース＋独自テキスト", mtype: "既製＋独自", mcost: "購入 67,815円〜（節目で買い足し）", good: "本格的・大会実績・中高まで", bad: "最高水準の費用・教材買い足し",
                 issue: "総額が最高水準、大会・上級志向", attack: "コスパ、大会志向でない子の受け皿" },
  edison:      { lesson: "90分×月2回", size: "1:4〜5", cur: "2年制、毎月1体ロボット製作", contest: "年1回程度の大会", fb: "教室ごと",
                 emp: "加盟教室が採用", wage: "要確認", train: "加盟時研修", ratio: "1:4〜5", trev: "しっかり見てもらえる",
                 mat: "アーテックロボ（自社製）", mtype: "独自", mcost: "購入 約3〜4.4万円", good: "安め・少人数・持ち帰り可", bad: "キット代・料金が非公開",
                 issue: "小3〜、2年で一巡、料金非公開", attack: "年中からの一貫性、長期カリキュラム" },
  proglab:     { lesson: "50/90分×月3回", size: "10名に講師3名", cur: "段階別4コース、最後に発表", contest: "プログラボカップ・WRO", fb: "見学可",
                 emp: "アルバイト（学生）", wage: "1,100〜2,000円", train: "最大3か月", ratio: "1:3〜4", trev: "熱心／学生講師でばらつき",
                 mat: "LEGO SPIKE・EV3・micro:bit", mtype: "既製", mcost: "貸与（持ち帰り不可）", good: "入会金・教材0円・駅近", bad: "上位コース高額・学生講師差",
                 issue: "LEGOを組む中心、持ち帰り不可、沿線偏在", attack: "身近な素材のものづくり、沿線外出店" },
  tamiya:      { lesson: "90分×月2回", size: "要確認", cur: "IchigoJamのテキスト入力／工作キット", contest: "自由製作コンテスト", fb: "教室ごと",
                 emp: "加盟教室が採用", wage: "要確認", train: "要確認", ratio: "要確認", trev: "口コミが少ない",
                 mat: "タミヤ工作キット＋IchigoJam", mtype: "既製", mcost: "購入 年1.3〜2万円", good: "最安クラス・本物の工作", bad: "口コミ少・教室が少ない",
                 issue: "教室数が少ない、低学年に弱い", attack: "低学年からの一貫性、内容の密度" },
  litalico:    { lesson: "90分×月4回〜", size: "1:1〜4", cur: "ゲーム・ロボ・3DCG等を個別計画で", contest: "年2回の発表会", fb: "個別指導計画・FB",
                 emp: "正社員中心＋アルバイト", wage: "1,230〜1,500円", train: "1on1・教室会議", ratio: "1:1〜4", trev: "質が高い／担当が変わる",
                 mat: "Scratch・LEGO・3Dプリンタ等", mtype: "既製＋独自", mcost: "月謝込み（教室備品）", good: "先生の質・個別対応", bad: "業界最高水準の料金・首都圏のみ",
                 issue: "月3万円、1都3県のみ", attack: "同じ少人数探究型を低価格・地域密着で" },
  techkids:    { lesson: "120分×月3回", size: "3〜6名にメンター1", cur: "Scratch→Swift/Unityでアプリ開発", contest: "Tech Kids Grand Prix", fb: "半年ごとの成果発表会",
                 emp: "大学生メンター", wage: "要確認", train: "選考・研修で認定", ratio: "1:3〜6", trev: "年が近く話しやすい",
                 mat: "QUREO・Swift・Unity", mtype: "独自＋既製", mcost: "教材費2,200円/月・PC持参", good: "本格度・発表文化", bad: "高額・渋谷のみ・受付停止中",
                 issue: "新規受付停止、渋谷のみ", attack: "本格志向層の受け皿" },
  qureo:       { lesson: "60分×月4回", size: "教室ごと（自習型）", cur: "オンライン教材で自学自習、検定準拠", contest: "プログラミング能力検定", fb: "保護者ページで進捗",
                 emp: "加盟塾の講師", wage: "1,300〜1,680円", train: "要確認", ratio: "教室ごと", trev: "当たり外れ・見守るだけ",
                 mat: "QUREO（オンライン教材）", mtype: "独自", mcost: "無料（月謝込み）", good: "安い・近所の塾・検定", bad: "教室差・自習中心・画面のみ",
                 issue: "画面内で完結、講師差", attack: "ハンズオン、対話型授業、小2未満" },
  zkai:        { lesson: "月1回分の教材（自宅）", size: "—", cur: "KOOVで組立→プログラミング（3年）", contest: "特になし", fb: "毎月の保護者ガイド",
                 emp: "講師なし", wage: "—", train: "—", ratio: "—", trev: "伴走者がいない",
                 mat: "KOOV（ソニー）", mtype: "既製", mcost: "キット代込み月額（所有）", good: "自分のペース・キット所有", bad: "飽きる・親の負担",
                 issue: "講座の縮小、講師不在", attack: "通信で続かない家庭の受け皿" },
  digitane:    { lesson: "受け放題（動画）", size: "—", cur: "マイクラ・Robloxでゲーム制作", contest: "要確認", fb: "要確認",
                 emp: "チャット対応", wage: "要確認", train: "要確認", ratio: "—", trev: "即時に解決しにくい",
                 mat: "動画教材・Minecraft・Roblox", mtype: "独自＋既製", mcost: "Minecraftは別途購入", good: "ゲームで夢中・安い", bad: "即時に質問できない・ゲーム目的化",
                 issue: "ゲーム依存、自走が前提", attack: "リアルな体験と講師の伴走" },
  codecamp:    { lesson: "90分（回数は要確認）", size: "オンライン5名程度", cur: "Scratch→ロボット→Unity", contest: "ジュニア・プログラミング検定", fb: "課題へのFB（オンライン）",
                 emp: "アルバイト", wage: "1,200〜1,500円", train: "研修あり", ratio: "要確認", trev: "丁寧／講師で差",
                 mat: "Scratch・toio・Sphero・Unity", mtype: "既製", mcost: "要確認", good: "教室がきれい・Unityまで", bad: "上位コース高額・講師差",
                 issue: "教室数が少ない、上位コース高額", attack: "幼児のものづくり、教室網の広さ" },
};

const SCALE = {
  steamcampus: "教室数不明（首都圏の商業施設中心）", human: "2,000教室以上・生徒2.7万人", robodone: "約100〜120教室（2022年）",
  crefus: "112教室＋海外", edison: "約850〜900教室（2021年頃）", proglab: "84校・約8,000名（2024年）", tamiya: "約100教室（古い数値）",
  litalico: "23拠点（1都3県）＋オンライン", techkids: "渋谷1校＋オンライン／新規受付停止", qureo: "3,000教室超（塾併設）",
  zkai: "通信教育（全国）", digitane: "オンライン＋提携教室", codecamp: "40教室以上＋オンライン",
};
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.title = "STEAM教育・子ども向けプログラミング教室 競合比較";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
const W = 13.333, M = 0.5;
let page = 0;

function base(title, kicker) {
  const s = pres.addSlide();
  page++;
  s.background = { color: C.white };
  if (kicker) s.addText(kicker, { x: M, y: 0.28, w: 8, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: C.yel, margin: 0, isTextBox: true });
  s.addText(title, { x: M, y: 0.55, w: W - 2 * M, h: 0.6, fontFace: FONT, fontSize: 24, bold: true, color: C.navy, margin: 0, isTextBox: true });
  s.addText(`出典：各社公式サイト・比較サイト・求人サイト等の検索結果要約（調査日 ${D.date}）。「要確認」「推定」を含む社内検討用資料`, {
    x: M, y: 7.08, w: 11, h: 0.25, fontFace: FONT, fontSize: 8, color: "8A92A3", margin: 0, isTextBox: true });
  s.addText(String(page), { x: W - M - 0.6, y: 7.08, w: 0.6, h: 0.25, fontFace: FONT, fontSize: 8, color: "8A92A3", align: "right", margin: 0, isTextBox: true });
  return s;
}

function catDot(c) { return CAT_COLOR[c.cat] || C.sub; }

function table(s, head, rows, opt) {
  const hdr = head.map(h => ({ text: h, options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center", valign: "middle" } }));
  const body = rows.map(r => r.cells.map((v, i) => {
    const o = { valign: "middle", color: C.ink };
    if (r.stemon) o.fill = { color: C.paleY };
    if (r.stemon && i === 0) o.bold = true;
    if (typeof v === "object" && v !== null) return { text: v.text, options: Object.assign(o, v.options || {}) };
    return { text: String(v), options: o };
  }));
  s.addTable([hdr, ...body], Object.assign({
    x: M, y: 1.35, w: W - 2 * M, fontFace: FONT, fontSize: 9, border: { type: "solid", pt: 0.5, color: C.line },
    margin: [2, 4, 2, 4], autoPage: false,
  }, opt));
}

function nameCell(c) {
  return { text: c.short, options: { color: c.key === "stemon" ? C.navy : C.ink, bold: true } };
}
function threatCell(t) {
  return { text: t, options: { bold: true, color: THREAT_COLOR[t] || C.sub, align: "center" } };
}

// =========== 1. 表紙 ===========
{
  const s = pres.addSlide(); page++;
  s.background = { color: C.navy };
  s.addShape(pres.shapes.OVAL, { x: 9.6, y: -1.2, w: 5.2, h: 5.2, fill: { color: C.yel, transparency: 15 }, line: { color: C.yel, transparency: 100 } });
  s.addShape(pres.shapes.OVAL, { x: 11.2, y: 4.6, w: 2.6, h: 2.6, fill: { color: "34426A" }, line: { color: "34426A" } });
  s.addText("社内戦略検討用", { x: 0.8, y: 1.6, w: 6, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: C.yel, margin: 0, isTextBox: true });
  s.addText("STEAM教育・子ども向け\nプログラミング教室 競合比較", { x: 0.8, y: 2.1, w: 9, h: 1.9, fontFace: FONT, fontSize: 38, bold: true, color: C.white, margin: 0, isTextBox: true, valign: "top" });
  s.addText("料金／レッスン内容／先生の質／使用教材／考えられる課題", { x: 0.8, y: 4.2, w: 9, h: 0.4, fontFace: FONT, fontSize: 16, color: "C9D2E6", margin: 0, isTextBox: true });
  s.addText(`ステモン＋競合13社（参考：ナリカ）｜調査日 ${D.date}`, { x: 0.8, y: 6.3, w: 9, h: 0.35, fontFace: FONT, fontSize: 12, color: "C9D2E6", margin: 0, isTextBox: true });
}

// =========== 2. エグゼクティブサマリー ===========
{
  const s = base("結論：価格は「中央帯」。勝ち筋は価格ではなく、講師品質と成果の見える化", "EXECUTIVE SUMMARY");
  const st = byKey.stemon;
  const cards = [
    { n: "01", h: "価格は通学型ロボット教室の中央帯", b: `ステモンの3年総額は約${man(st.y3)}万円（試算）。ヒューマンアカデミー・プログラボ・エジソンアカデミーとほぼ同水準。高価格帯はLITALICO（約${man(byKey.litalico.y3)}万円）・クレファス（約${man(byKey.crefus.y3)}万円）、低価格帯はオンライン系（18〜23万円）。` },
    { n: "02", h: "競合の共通の弱点＝ステモンの攻めどころ", b: "①講師品質のばらつき（業務委託・学生講師・加盟塾依存）\n②教材の買い足し・総額が見えにくい\n③作品を持ち帰れない／成長が見えにくい\n→ 教材貸出・少人数・年中からの一貫性が効く" },
    { n: "03", h: "市場は再編期。流出層の受け皿チャンス", b: "STEAM Campus（LEGOブランド喪失）、Tech Kids School（新規受付停止）、Z会（講座縮小）。一方でヒューマンアカデミー（2,000教室超）とQUREO（3,000教室超）の規模・塾併設FCが最大の脅威。" },
  ];
  cards.forEach((c, i) => {
    const x = M + i * 4.16;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.45, w: 3.95, h: 3.9, fill: { color: C.pale }, line: { color: C.pale }, rectRadius: 0.12 });
    s.addText(c.n, { x: x + 0.25, y: 1.6, w: 1, h: 0.55, fontFace: FONT, fontSize: 26, bold: true, color: C.yel, margin: 0, isTextBox: true });
    s.addText(c.h, { x: x + 0.25, y: 2.15, w: 3.5, h: 0.75, fontFace: FONT, fontSize: 15, bold: true, color: C.navy, margin: 0, isTextBox: true, valign: "top" });
    s.addText(c.b, { x: x + 0.25, y: 2.95, w: 3.5, h: 2.3, fontFace: FONT, fontSize: 12, color: C.ink, margin: 0, isTextBox: true, valign: "top", lineSpacingMultiple: 1.2 });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 5.6, w: W - 2 * M, h: 1.25, fill: { color: C.navy }, line: { color: C.navy }, rectRadius: 0.12 });
  s.addText("優先アクション", { x: M + 0.3, y: 5.75, w: 2.2, h: 0.95, fontFace: FONT, fontSize: 15, bold: true, color: C.yel, margin: 0, isTextBox: true, valign: "middle" });
  s.addText([
    { text: "① 「3年総額」を公式に提示し、競合との総額比較で説明する", options: { breakLine: true } },
    { text: "② 作品持ち帰り＋毎回の保護者レポートで「成長の見える化」を標準化", options: { breakLine: true } },
    { text: "③ 小4〜6の出口（テキストコーディング・外部大会挑戦）を強化し高学年の退会を防ぐ" },
  ], { x: M + 2.6, y: 5.72, w: 9.6, h: 1.0, fontFace: FONT, fontSize: 12, color: C.white, margin: 0, isTextBox: true, valign: "middle", paraSpaceAfter: 2 });
}

// =========== 3. 調査概要 ===========
{
  const s = base("調査の概要と数値の読み方", "SCOPE & METHOD");
  const items = [
    ["目的", "社内の戦略検討。ステモンの立ち位置と打ち手を明らかにする"],
    ["対象", "ステモン＋競合13社（ロボット・ものづくり系7／プログラミング特化3／オンライン・通信3）。参考：ナリカ（教材メーカー）"],
    ["比較軸", "料金（1年目・3年総額）／レッスン内容／先生の質（募集要項・口コミ）／使用教材／口コミ／課題と攻めどころ"],
    ["方法", "Web検索結果（公式サイト・比較サイト・口コミサイト・求人サイトの要約）から収集"],
    ["総額の計算", "入会金＋（月謝＋月額諸費用）×12か月＋教材・その他。各社の前提は元データExcelの「料金比較」シートに記載"],
  ];
  items.forEach((it, i) => {
    const y = 1.45 + i * 0.78;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: 1.7, h: 0.6, fill: { color: C.navy }, line: { color: C.navy }, rectRadius: 0.08 });
    s.addText(it[0], { x: M, y, w: 1.7, h: 0.6, fontFace: FONT, fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(it[1], { x: M + 1.9, y, w: 6.2, h: 0.6, fontFace: FONT, fontSize: 11, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
  });
  // 注意ボックス
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.9, y: 1.45, w: 3.93, h: 3.7, fill: { color: "FDECEC" }, line: { color: "FDECEC" }, rectRadius: 0.12 });
  s.addText("注意", { x: 9.15, y: 1.6, w: 3.4, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: C.hi, margin: 0, isTextBox: true });
  s.addText("調査環境の制約により、各社公式ページの本文は直接閲覧できていません。数値は検索結果の要約に基づきます。\n\n社外利用・最終判断の前に、元データExcelの「要確認リスト」（17項目）を公式ページで確認してください。", {
    x: 9.15, y: 2.05, w: 3.45, h: 3.0, fontFace: FONT, fontSize: 11, color: C.ink, margin: 0, isTextBox: true, valign: "top" });
  // 確度凡例
  const leg = [["公式", "運営会社・公式サイト由来"], ["二次", "比較・口コミ・ブログ由来"], ["推定", "本資料での試算"], ["要確認", "古い・食い違い・未確認"]];
  s.addText("確度の表記", { x: M, y: 5.5, w: 3, h: 0.35, fontFace: FONT, fontSize: 12, bold: true, color: C.navy, margin: 0, isTextBox: true });
  leg.forEach((l, i) => {
    const x = M + i * 3.08;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 5.95, w: 2.9, h: 0.75, fill: { color: C.pale }, line: { color: C.pale }, rectRadius: 0.08 });
    s.addText([{ text: l[0], options: { bold: true, color: C.navy, fontSize: 12, breakLine: true } }, { text: l[1], options: { fontSize: 10, color: C.sub } }],
      { x: x + 0.15, y: 5.98, w: 2.65, h: 0.7, fontFace: FONT, margin: 0, isTextBox: true, valign: "middle" });
  });
}

// =========== 4. 比較対象 ===========
{
  const s = base("比較対象：3分類・13社＋参考1社", "COMPETITORS");
  const groups = [
    ["ロボット・ものづくり系", C.robo, COS.filter(c => c.cat === "ロボット・ものづくり")],
    ["プログラミング特化系", C.prog, COS.filter(c => c.cat === "プログラミング特化")],
    ["オンライン・通信系", C.online, COS.filter(c => c.cat === "オンライン・通信")],
  ];
  const colW = 3.0, gap = 0.18;
  groups.forEach((g, gi) => {
    const x = M + gi * (colW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.4, w: colW, h: 0.5, fill: { color: g[1] }, line: { color: g[1] }, rectRadius: 0.08 });
    s.addText(g[0], { x, y: 1.4, w: colW, h: 0.5, fontFace: FONT, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
    g[2].forEach((c, i) => {
      const y = 2.02 + i * 0.7;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: colW, h: 0.62, fill: { color: C.pale }, line: { color: C.pale }, rectRadius: 0.08 });
      s.addText([{ text: c.short, options: { bold: true, fontSize: 12, color: C.ink, breakLine: true } }, { text: SCALE[c.key], options: { fontSize: 9, color: C.sub } }],
        { x: x + 0.15, y, w: colW - 0.8, h: 0.62, fontFace: FONT, margin: 0, isTextBox: true, valign: "middle" });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + colW - 0.58, y: y + 0.15, w: 0.46, h: 0.3, fill: { color: THREAT_COLOR[c.threat] }, line: { color: THREAT_COLOR[c.threat] }, rectRadius: 0.05 });
      s.addText(c.threat, { x: x + colW - 0.58, y: y + 0.15, w: 0.46, h: 0.3, fontFace: FONT, fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
    });
  });
  // 参考
  const x = M + 3 * (colW + gap);
  const w = W - M - x;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.4, w, h: 0.5, fill: { color: C.sub }, line: { color: C.sub }, rectRadius: 0.08 });
  s.addText("参考（競合ではない）", { x, y: 1.4, w, h: 0.5, fontFace: FONT, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.02, w, h: 2.4, fill: { color: C.pale }, line: { color: C.pale }, rectRadius: 0.08 });
  s.addText([
    { text: "ナリカ", options: { bold: true, fontSize: 12, color: C.ink, breakLine: true } },
    { text: "理科・プログラミング教材メーカー（LEGO Education正規代理店）。子ども向け教室は運営していない。→ 教材調達・学校連携のパートナー候補", options: { fontSize: 10, color: C.ink } },
  ], { x: x + 0.15, y: 2.1, w: w - 0.3, h: 2.25, fontFace: FONT, margin: 0, isTextBox: true, valign: "top" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.55, w, h: 1.55, fill: { color: C.pale }, line: { color: C.pale }, rectRadius: 0.08 });
  s.addText([
    { text: "今後の注視対象", options: { bold: true, fontSize: 11, color: C.ink, breakLine: true } },
    { text: "KOOVパートナー教室（ソニー）、学研 もののしくみ研究室", options: { fontSize: 10, color: C.ink } },
  ], { x: x + 0.15, y: 4.62, w: w - 0.3, h: 1.4, fontFace: FONT, margin: 0, isTextBox: true, valign: "top" });
  // 凡例
  s.addText("右の色ラベル＝ステモンにとっての脅威度（高・中・低／定性評価）", { x: x, y: 6.25, w: w, h: 0.6, fontFace: FONT, fontSize: 9, color: C.sub, margin: 0, isTextBox: true });
}

// =========== 5. 基本情報 ===========
{
  const s = base("基本情報：規模はヒューマンアカデミーとQUREOが突出", "① BASIC INFO");
  table(s, ["社名", "運営会社", "対象", "規模", "運営形態", "脅威度"],
    COS.map(c => ({ stemon: c.key === "stemon", cells: [nameCell(c), c.operator.split("（")[0].split("。")[0], c.target.split("（")[0], c.scale, c.form.split("（")[0], threatCell(c.threat)] })),
    { colW: [1.8, 2.9, 1.7, 3.6, 1.6, 0.73], rowH: 0.36 });
}

// =========== 6. 料金① 3年総額 ===========
{
  const s = base("料金①：3年総額で見ると、ステモンは通学型の中央帯（約46.5万円）", "② PRICE");
  const sorted = [...COS].sort((a, b) => b.y3 - a.y3);
  // Googleスライド変換でネイティブグラフのラベルが崩れるため、図形で描画する
  s.addText("3年継続の総額（試算、万円）", { x: M, y: 1.35, w: 7.9, h: 0.3, fontFace: FONT, fontSize: 12, bold: true, color: C.navy, align: "center", margin: 0, isTextBox: true });
  const lx = M, lw = 1.75, bx = lx + lw + 0.1, bwMax = 7.9 - lw - 0.1 - 0.85;
  const top = 1.8, rowH = 0.36, barH = 0.25, maxV = sorted[0].y3;
  sorted.forEach((c, i) => {
    const y = top + i * rowH, st = c.key === "stemon";
    const bw = bwMax * c.y3 / maxV;
    s.addText(c.short, { x: lx, y, w: lw, h: barH, fontFace: FONT, fontSize: 10, bold: st, color: C.ink, align: "right", valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y, w: bw, h: barH, fill: { color: st ? C.yel : C.navy }, line: { color: st ? C.yel : C.navy } });
    s.addText(`${man(c.y3)}万`, { x: bx + bw + 0.08, y, w: 0.8, h: barH, fontFace: FONT, fontSize: 10, bold: st, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
  });
  const x = 8.75, w = W - M - x;
  const pts = [
    ["中央帯に密集", "ステモン・エジソン・ロボ団・ヒューマン・プログラボ・STEAM Campusが44〜52万円に集中。価格での差別化は難しい"],
    ["高価格帯", "LITALICO（月約3万円）、Tech Kids（受付停止）、クレファス（教材買い足し）"],
    ["低価格帯", "タミヤ・QUREO（月1万円以下）、Z会・デジタネ（オンライン）。講師の伴走がない／薄い"],
    ["見えない費用", "クレファスの追加キット、エジソンの3年目教材、ロボ団のタブレット代などは未計上＝実際はさらに高い"],
  ];
  pts.forEach((p, i) => {
    const y = 1.45 + i * 1.37;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 1.22, fill: { color: i === 0 ? C.paleY : C.pale }, line: { color: i === 0 ? C.paleY : C.pale }, rectRadius: 0.08 });
    s.addText([{ text: p[0], options: { bold: true, fontSize: 12, color: C.navy, breakLine: true } }, { text: p[1], options: { fontSize: 10, color: C.ink } }],
      { x: x + 0.15, y: y + 0.05, w: w - 0.3, h: 1.12, fontFace: FONT, margin: 0, isTextBox: true, valign: "middle" });
  });
}

// =========== 7. 料金② 内訳 ===========
{
  const s = base("料金②：内訳と1時間あたり単価。ステモンは約3,500円/時で「中の下」", "② PRICE");
  const yen = v => v === 0 ? "0" : v.toLocaleString("ja-JP");
  const unit = c => c.min_per_month ? Math.round((c.m1 + c.misc) / (c.min_per_month / 60)).toLocaleString("ja-JP") : "—";
  table(s, ["社名", "入会金", "月謝（1年目）", "教材・その他（1年目）", "1年目総額", "3年総額", "1時間単価", "計算の前提（要約）"],
    COS.map(c => ({ stemon: c.key === "stemon", cells: [nameCell(c), yen(c.entry), yen(c.m1 + c.misc), yen(c.mat1), { text: `${man(c.y1)}万`, options: { align: "right" } }, { text: `${man(c.y3)}万`, options: { align: "right", bold: true } }, { text: unit(c), options: { align: "right" } },
      { text: c.price_assume.length > 58 ? c.price_assume.slice(0, 57) + "…" : c.price_assume, options: { fontSize: 8 } }] })),
    { colW: [1.8, 0.85, 1.05, 1.25, 0.95, 0.95, 0.9, 4.58], rowH: 0.36 });
  s.addText("月謝は月額諸費用込み。1時間単価＝月額÷月の授業時間（不明・通信型は—）。前提の全文は元データExcel「料金比較」シートC列", { x: M, y: 6.8, w: W - 2 * M, h: 0.24, fontFace: FONT, fontSize: 8, color: C.sub, margin: 0, isTextBox: true });
}

// =========== 8. レッスン内容 ===========
{
  const s = base("レッスン内容：月の授業回数・人数・大会との接続が差になる", "③ LESSON");
  table(s, ["社名", "時間・回数", "クラス人数", "カリキュラムの特徴", "大会・検定", "保護者への共有"],
    COS.map(c => ({ stemon: c.key === "stemon", cells: [nameCell(c), S[c.key].lesson, S[c.key].size, S[c.key].cur, S[c.key].contest, S[c.key].fb] })),
    { colW: [1.8, 2.0, 1.6, 3.4, 2.1, 1.43], rowH: 0.36 });
}

// =========== 9. 先生の質 ===========
{
  const s = base("先生の質：多くがアルバイト・業務委託で「教室差」が口コミの定番不満", "④ TEACHER");
  table(s, ["社名", "雇用形態", "時給・給与", "研修", "講師1人あたり", "口コミ傾向（良い／不満）"],
    COS.map(c => ({ stemon: c.key === "stemon", cells: [nameCell(c), S[c.key].emp, S[c.key].wage, S[c.key].train, S[c.key].ratio, S[c.key].trev] })),
    { colW: [1.8, 1.9, 2.45, 1.9, 1.3, 2.98], rowH: 0.36 });
}

// =========== 10. 使用教材 ===========
{
  const s = base("使用教材：「購入型」「貸与型」「月謝込み型」の3類型", "⑤ MATERIALS");
  table(s, ["社名", "使用教材", "区分", "費用負担"],
    COS.map(c => ({ stemon: c.key === "stemon", cells: [nameCell(c), S[c.key].mat, S[c.key].mtype, S[c.key].mcost] })),
    { x: M, w: 8.3, colW: [1.8, 2.8, 1.0, 2.7], rowH: 0.36 });
  const x = 9.05, w = W - M - x;
  const t = [
    ["購入型", "ヒューマン・クレファス・エジソン・タミヤ・STEAM Campus", "作品を持ち帰れるが、初期費用・買い足しが不満の定番"],
    ["貸与型", "ステモン・ロボ団・プログラボ", "初期費用は低いが、ロボ団・プログラボは持ち帰り不可が不満"],
    ["月謝込み型", "LITALICO・QUREO", "総額は分かりやすいが、月謝に上乗せされる"],
  ];
  t.forEach((r, i) => {
    const y = 1.35 + i * 1.62;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 1.48, fill: { color: C.pale }, line: { color: C.pale }, rectRadius: 0.08 });
    s.addText([{ text: r[0], options: { bold: true, fontSize: 13, color: C.navy, breakLine: true } }, { text: r[1], options: { fontSize: 9, color: C.sub, breakLine: true } }, { text: r[2], options: { fontSize: 10, color: C.ink } }],
      { x: x + 0.15, y: y + 0.05, w: w - 0.3, h: 1.38, fontFace: FONT, margin: 0, isTextBox: true, valign: "middle" });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 6.2, w, h: 0.7, fill: { color: C.paleY }, line: { color: C.paleY }, rectRadius: 0.08 });
  s.addText("ステモンは「貸与で初期費用が低い」＋「成果を家に持ち帰れる工夫」で両取りを狙える", { x: x + 0.15, y: 6.2, w: w - 0.3, h: 0.7, fontFace: FONT, fontSize: 10, bold: true, color: C.navy, margin: 0, isTextBox: true, valign: "middle" });
}

// =========== 11. 口コミ ===========
{
  const s = base("口コミ傾向：不満は「費用」「講師差」「成果が見えない」に集中", "⑥ REVIEWS");
  table(s, ["社名", "褒められている点", "不満が多い点"],
    COS.map(c => ({ stemon: c.key === "stemon", cells: [nameCell(c), S[c.key].good, S[c.key].bad] })),
    { colW: [1.8, 5.07, 5.46], rowH: 0.36, fontSize: 10 });
}

// =========== 12. 課題と攻めどころ ===========
{
  const s = base("各社の課題と、ステモンが攻められるポイント", "⑦ ISSUES");
  const comp = COS.filter(c => c.key !== "stemon");
  table(s, ["社名", "考えられる課題（弱み）", "ステモンの攻めどころ", "脅威度"],
    comp.map(c => ({ cells: [nameCell(c), S[c.key].issue, S[c.key].attack, threatCell(c.threat)] })),
    { colW: [1.8, 4.9, 4.83, 0.8], rowH: 0.39, fontSize: 10 });
}

// =========== 13. ポジショニングマップ ===========
{
  const s = base("ポジショニング：「40〜55万円×ものづくり寄り」が激戦区", "⑧ POSITIONING");
  const px = 1.3, py = 1.45, pw = 8.1, ph = 4.85; // plot area
  const xmin = 10, xmax = 115, ymin = 0.5, ymax = 5.5;
  const X = v => px + (v - xmin) / (xmax - xmin) * pw;
  const Y = v => py + (ymax - v) / (ymax - ymin) * ph;
  // 激戦区
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: X(40), y: Y(5.35), w: X(57) - X(40), h: Y(3.1) - Y(5.35), fill: { color: C.paleY }, line: { color: C.yel, width: 1, dashType: "dash" }, rectRadius: 0.1 });
  s.addText("激戦区", { x: X(57) - 0.75, y: Y(3.1) - 0.3, w: 1.0, h: 0.26, fontFace: FONT, fontSize: 9, bold: true, color: C.mid, margin: 0, isTextBox: true });
  // axes
  s.addShape(pres.shapes.LINE, { x: px, y: py + ph, w: pw, h: 0, line: { color: C.sub, width: 1 } });
  s.addShape(pres.shapes.LINE, { x: px, y: py, w: 0, h: ph, line: { color: C.sub, width: 1 } });
  [20, 40, 60, 80, 100].forEach(v => {
    s.addText(`${v}万`, { x: X(v) - 0.4, y: py + ph + 0.05, w: 0.8, h: 0.25, fontFace: FONT, fontSize: 9, color: C.sub, align: "center", margin: 0, isTextBox: true });
  });
  s.addText("3年総額（試算）→ 高い", { x: px + pw - 3, y: py + ph + 0.32, w: 3, h: 0.28, fontFace: FONT, fontSize: 10, bold: true, color: C.sub, align: "right", margin: 0, isTextBox: true });
  s.addText("ものづくり・実体験", { x: 0.3, y: py - 0.05, w: 1.0, h: 0.6, fontFace: FONT, fontSize: 9, bold: true, color: C.sub, align: "center", margin: 0, isTextBox: true });
  s.addText("画面内の\nプログラミング", { x: 0.3, y: py + ph - 0.6, w: 1.0, h: 0.6, fontFace: FONT, fontSize: 9, bold: true, color: C.sub, align: "center", margin: 0, isTextBox: true });
  // label offsets (inch): [dx, dy, align]
  const OFF = {
    stemon: [-1.62, -0.13, "right"], steamcampus: [0.2, -0.13, "left"], human: [0.2, -0.13, "left"], edison: [-1.6, -0.13, "right"],
    robodone: [-1.55, -0.08, "right"], proglab: [0.2, -0.1, "left"], crefus: [0.2, -0.13, "left"], tamiya: [-1.55, -0.13, "right"],
    litalico: [-1.6, -0.13, "right"], techkids: [0.2, -0.13, "left"], qureo: [0.2, -0.13, "left"], zkai: [0.2, -0.13, "left"],
    digitane: [0.2, -0.13, "left"], codecamp: [0.2, -0.13, "left"],
  };
  const JIT = { human: 0.12, edison: -0.12, robodone: -0.1, proglab: 0.1 };
  COS.forEach(c => {
    const v = c.y3 / 10000, sc = c.exp_score + (JIT[c.key] || 0);
    const r = c.key === "stemon" ? 0.2 : 0.13;
    const col = catDot(c);
    s.addShape(pres.shapes.OVAL, { x: X(v) - r, y: Y(sc) - r, w: 2 * r, h: 2 * r, fill: { color: col }, line: { color: C.white, width: 1 } });
    const o = OFF[c.key];
    s.addText(c.short, { x: X(v) + o[0], y: Y(sc) + o[1], w: 1.35, h: 0.26, fontFace: FONT, fontSize: c.key === "stemon" ? 11 : 9, bold: c.key === "stemon", color: C.ink, align: o[2], margin: 0, isTextBox: true });
  });
  // 凡例
  const legend = [["自社", C.yel], ["ロボット・ものづくり", C.robo], ["プログラミング特化", C.prog], ["オンライン・通信", C.online]];
  legend.forEach((l, i) => {
    s.addShape(pres.shapes.OVAL, { x: 1.4 + i * 1.75, y: 6.84, w: 0.14, h: 0.14, fill: { color: l[1] }, line: { color: l[1] } });
    s.addText(l[0], { x: 1.6 + i * 1.75, y: 6.78, w: 1.6, h: 0.26, fontFace: FONT, fontSize: 9, color: C.ink, margin: 0, isTextBox: true });
  });
  const x = 9.85, w = W - M - x;
  const pts = [
    ["読み取り", "価格帯が同じ競合が多く、価格では差がつかない。同じ価格帯の中で「何が違うか」を説明できることが勝敗を分ける"],
    ["ステモンの独自領域", "年中から・物理の仕組み・教材貸出・少人数。「ものづくり×探究」の最上位に位置"],
    ["空白", "「ものづくり寄り×30万円台」はタミヤのみ。低価格の入門コース／短期講座で入口を作る余地"],
  ];
  pts.forEach((p, i) => {
    const y = 1.45 + i * 1.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 1.6, fill: { color: i === 1 ? C.paleY : C.pale }, line: { color: i === 1 ? C.paleY : C.pale }, rectRadius: 0.08 });
    s.addText([{ text: p[0], options: { bold: true, fontSize: 12, color: C.navy, breakLine: true } }, { text: p[1], options: { fontSize: 10, color: C.ink } }],
      { x: x + 0.12, y: y + 0.05, w: w - 0.24, h: 1.5, fontFace: FONT, margin: 0, isTextBox: true, valign: "middle" });
  });
  s.addText("縦軸は定性評価（元データの「ポジショニング」シートで調整可）", { x: x, y: 6.75, w: w, h: 0.25, fontFace: FONT, fontSize: 8, color: C.sub, margin: 0, isTextBox: true });
}

// =========== 14. SWOT ===========
{
  const s = base("ステモンのSWOT（競合比較から導出）", "⑨ SWOT");
  const Q = [["S", "強み", C.robo, D.swot.S], ["W", "弱み", C.hi, D.swot.W], ["O", "機会", C.online, D.swot.O], ["T", "脅威", C.mid, D.swot.T]];
  const bw = (W - 2 * M - 0.2) / 2, bh = 2.72;
  Q.forEach((q, i) => {
    const x = M + (i % 2) * (bw + 0.2), y = 1.35 + Math.floor(i / 2) * (bh + 0.15);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: bw, h: bh, fill: { color: C.pale }, line: { color: C.pale }, rectRadius: 0.1 });
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.18, w: 0.55, h: 0.55, fill: { color: q[2] }, line: { color: q[2] } });
    s.addText(q[0], { x: x + 0.2, y: y + 0.18, w: 0.55, h: 0.55, fontFace: FONT, fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(q[1], { x: x + 0.9, y: y + 0.2, w: 3, h: 0.5, fontFace: FONT, fontSize: 15, bold: true, color: C.navy, valign: "middle", margin: 0, isTextBox: true });
    s.addText(q[3].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < q[3].length - 1 } })),
      { x: x + 0.25, y: y + 0.82, w: bw - 0.45, h: bh - 0.95, fontFace: FONT, fontSize: 12, color: C.ink, margin: 0, isTextBox: true, valign: "top", paraSpaceAfter: 5 });
  });
}

// =========== 15. 提言 ===========
{
  const s = base("提言：7つの打ち手（優先度・期間つき）", "⑩ ACTIONS");
  const PR = { "高": C.hi, "中": C.mid, "低": C.lo };
  D.actions.forEach((a, i) => {
    const y = 1.35 + i * 0.79;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y, w: W - 2 * M, h: 0.68, fill: { color: a[2] === "高" ? C.paleY : C.pale }, line: { color: a[2] === "高" ? C.paleY : C.pale }, rectRadius: 0.08 });
    s.addText(a[0], { x: M + 0.2, y, w: 1.9, h: 0.68, fontFace: FONT, fontSize: 12, bold: true, color: C.navy, valign: "middle", margin: 0, isTextBox: true });
    s.addText(a[1], { x: M + 2.1, y, w: 7.9, h: 0.68, fontFace: FONT, fontSize: 11, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M + 10.2, y: y + 0.17, w: 0.7, h: 0.34, fill: { color: PR[a[2]] }, line: { color: PR[a[2]] }, rectRadius: 0.05 });
    s.addText(`優先${a[2]}`, { x: M + 10.2, y: y + 0.17, w: 0.7, h: 0.34, fontFace: FONT, fontSize: 9, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(a[3], { x: M + 11.0, y, w: 1.3, h: 0.68, fontFace: FONT, fontSize: 10, color: C.sub, align: "center", valign: "middle", margin: 0, isTextBox: true });
  });
}

// =========== 16. 要確認・次のステップ ===========
{
  const s = base("要確認事項と次のステップ", "NEXT STEPS");
  s.addText("優先して確認すべき数値", { x: M, y: 1.35, w: 6, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: C.navy, margin: 0, isTextBox: true });
  const chk = [
    "ステモン自社：入会金（16,500円 / 5,500円）・施設管理費（880円 / 2,200円）の正、小4〜の月謝",
    "ロボ団：コース別の月謝、教材レンタル料・タブレット代の有無",
    "クレファス：コース別の月謝（3年総額が別サイトの試算と大きく食い違う）",
    "STEAM Campus：コース別の月謝・設備費・月の回数",
    "QUREO：入会金と教室ごとの月謝（近隣の加盟塾で確認）",
    "Tech Kids School：新規受付停止の時期・理由・再開予定",
  ];
  s.addText(chk.map((t, i) => ({ text: t, options: { bullet: { type: "number" }, breakLine: i < chk.length - 1 } })),
    { x: M, y: 1.85, w: 6.9, h: 4.2, fontFace: FONT, fontSize: 11, color: C.ink, margin: 0, isTextBox: true, valign: "top", paraSpaceAfter: 6 });
  s.addText("全17項目は元データExcelの「要確認リスト」シート", { x: M, y: 6.3, w: 6.9, h: 0.3, fontFace: FONT, fontSize: 9, color: C.sub, margin: 0, isTextBox: true });
  const x = 7.8, w = W - M - x;
  const steps = [["1", "数値の確定", "要確認リストを公式ページ・社内料金表で確認（〜2週間）"], ["2", "エリア別の競合マップ", "主要教室の半径2km内の競合教室を地図化し、出店・集客の優先度を決める"], ["3", "施策の具体化", "3年総額の提示と保護者レポートの試行（直営校から）"], ["4", "定期更新", "半年ごとに料金・教室数を更新（Excelの青字セルを書き換え）"]];
  steps.forEach((st, i) => {
    const y = 1.35 + i * 1.3;
    s.addShape(pres.shapes.OVAL, { x, y: y + 0.1, w: 0.6, h: 0.6, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(st[0], { x, y: y + 0.1, w: 0.6, h: 0.6, fontFace: FONT, fontSize: 16, bold: true, color: C.yel, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText([{ text: st[1], options: { bold: true, fontSize: 13, color: C.navy, breakLine: true } }, { text: st[2], options: { fontSize: 10, color: C.ink } }],
      { x: x + 0.8, y, w: w - 0.8, h: 1.1, fontFace: FONT, margin: 0, isTextBox: true, valign: "top" });
  });
}

// =========== 17. 参考：ナリカ ===========
{
  const s = base("参考：ナリカは競合ではなく「教材サプライヤー」", "APPENDIX");
  table(s, ["項目", "内容"], D.narika.rows.map(r => ({ cells: [{ text: r[0], options: { bold: true } }, r[1]] })),
    { colW: [2.2, 10.13], rowH: 0.62, fontSize: 11 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: M, y: 6.0, w: W - 2 * M, h: 0.8, fill: { color: C.paleY }, line: { color: C.paleY }, rectRadius: 0.08 });
  s.addText("示唆：理科実験・LEGO／micro:bit教材の調達、教員研修（NSA）ノウハウ、学校向け販路での協業余地を検討", { x: M + 0.2, y: 6.0, w: W - 2 * M - 0.4, h: 0.8, fontFace: FONT, fontSize: 12, bold: true, color: C.navy, valign: "middle", margin: 0, isTextBox: true });
}

pres.writeFile({ fileName: OUT }).then(f => console.log("saved", f));
