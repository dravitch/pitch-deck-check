/* Pitch deck checks against the published PSDAIC pitch-training rules (Nov. 19, 2026).
   Runs entirely in the browser: the file is never uploaded. No AI. Pure function of the .pptx content. */
(function (root) {
  const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const ACRO_OK = new Set(["OK", "AI", "US", "UK", "TV", "PM", "AM", "ID", "I", "A", "II", "III", "IV", "CEO", "Q&A"]);
  const OUTLINE = /^(agenda|outline|overview|table of contents|contents|plan|sommaire|ordre du jour|table des mati[eè]res)$/i;
  const PERSONAL_HOSTS = /(claude\.ai|chatgpt\.com|chat\.openai\.com|gemini\.google\.com|localhost|127\.0\.0\.1|drive\.google\.com|docs\.google\.com|dropbox\.com|onedrive\.live\.com|1drv\.ms)/i;

  function num(name) { const m = name.match(/(\d+)\.xml$/); return m ? +m[1] : 0; }
  function words(t) { return (t.match(/\S+/g) || []).length; }

  async function analyse(zip, parseXML) {
    const slideNames = Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a, b) => num(a) - num(b));
    const slides = [];
    for (const name of slideNames) {
      const doc = parseXML(await zip.file(name).async("string"));
      const relName = name.replace("slides/", "slides/_rels/") + ".rels";
      const rels = {};
      if (zip.file(relName)) {
        const rd = parseXML(await zip.file(relName).async("string"));
        for (const r of Array.from(rd.getElementsByTagName("Relationship")))
          rels[r.getAttribute("Id")] = { target: r.getAttribute("Target"), type: r.getAttribute("Type"), ext: r.getAttribute("TargetMode") === "External" };
      }
      // text runs and sizes
      const runs = []; let unknownSize = 0;
      for (const r of Array.from(doc.getElementsByTagNameNS(NS_A, "r"))) {
        const t = Array.from(r.getElementsByTagNameNS(NS_A, "t")).map(x => x.textContent).join("");
        if (!t.trim()) continue;
        const rPr = r.getElementsByTagNameNS(NS_A, "rPr")[0];
        const sz = rPr && rPr.getAttribute("sz") ? +rPr.getAttribute("sz") / 100 : null;
        if (sz === null) unknownSize++;
        runs.push({ t, sz });
      }
      const paras = Array.from(doc.getElementsByTagNameNS(NS_A, "p")).map(p => Array.from(p.getElementsByTagNameNS(NS_A, "t")).map(x => x.textContent).join("").trim()).filter(Boolean);
      const text = paras.join("\n");
      // title
      let title = "";
      for (const sp of Array.from(doc.getElementsByTagNameNS(NS_P, "sp"))) {
        const ph = sp.getElementsByTagNameNS(NS_P, "ph")[0];
        if (ph && /title|ctrTitle/.test(ph.getAttribute("type") || "")) { title = Array.from(sp.getElementsByTagNameNS(NS_A, "t")).map(x => x.textContent).join(" ").trim(); break; }
      }
      if (!title) title = paras[0] || "";
      // pictures and alt text
      const pics = Array.from(doc.getElementsByTagNameNS(NS_P, "pic"));
      const noAlt = pics.filter(p => { const c = p.getElementsByTagNameNS(NS_P, "cNvPr")[0]; return !c || !(c.getAttribute("descr") || "").trim(); }).length;
      // links
      const links = [];
      for (const h of Array.from(doc.getElementsByTagNameNS(NS_A, "hlinkClick"))) {
        const id = h.getAttributeNS(NS_R, "id") || h.getAttribute("r:id"); const r = rels[id];
        if (r && r.ext) links.push(r.target);
      }
      for (const m of text.match(/\b(?:https?:\/\/|www\.)?[a-z0-9.-]+\.(?:ai|com|ca|org|net|io|gc\.ca|app|dev)(?:\/[^\s)]*)?/gi) || [])
        if (/\/|www\.|\.[a-z]{2,}\//i.test(m) || PERSONAL_HOSTS.test(m)) links.push(m);
      // notes
      let notes = "";
      const notesRel = Object.values(rels).find(r => /notesSlide$/.test(r.type));
      if (notesRel) {
        const nn = "ppt/" + notesRel.target.replace(/^\.\.\//, "");
        if (zip.file(nn)) {
          const nd = parseXML(await zip.file(nn).async("string"));
          notes = Array.from(nd.getElementsByTagNameNS(NS_A, "p")).map(p => Array.from(p.getElementsByTagNameNS(NS_A, "t")).map(x => x.textContent).join("")).join("\n")
            .replace(/^\s*\d+\s*$/gm, "").trim();
        }
      }
      slides.push({ n: slides.length + 1, title, text, runs, unknownSize, pics: pics.length, noAlt, links: [...new Set(links)], notes });
    }
    return check(slides);
  }

  function acronyms(s) {
    // Lines written entirely in capitals are headings: their words are not acronyms.
    s = s.split(/\r?\n/).filter(l => /[a-zà-ÿ]/.test(l)).join("\n");
    return (s.match(/\b[A-Z][A-Z0-9]*(?:[\/&][A-Z0-9]+)*\b/g) || []).filter(w => /[A-Z].*[A-Z]|[A-Z]\/[A-Z]/.test(w) && !ACRO_OK.has(w) && !/^[IVX]+$/.test(w));
  }
  function timeOf(s) { const m = s.match(/(\d{1,2}):(\d{2})/); return m ? +m[1] * 60 + +m[2] : null; }

  function check(slides) {
    const R = []; const add = (id, rule, verdict, detail, where) => R.push({ id, rule, verdict, detail, where: where || "" });
    // 1. font size
    const small = slides.map(s => ({ n: s.n, k: s.runs.filter(r => r.sz !== null && r.sz < 18).length, min: Math.min(...s.runs.filter(r => r.sz !== null).map(r => r.sz), Infinity), unk: s.unknownSize })).filter(x => x.k);
    const totalUnknown = slides.reduce((a, s) => a + s.unknownSize, 0);
    add("A1", "Text 18 pt or larger", small.length ? "Does not fit" : (totalUnknown ? "To check" : "Fits"),
      small.length ? `${small.reduce((a, x) => a + x.k, 0)} text runs below 18 pt on ${small.length} slide(s); smallest ${Math.min(...small.map(x => x.min))} pt.` + (totalUnknown ? ` ${totalUnknown} runs inherit their size from the layout and were not measured.` : "")
        : (totalUnknown ? `${totalUnknown} runs inherit their size from the layout: check them on screen.` : "All measured text is 18 pt or larger."),
      small.map(x => `slide ${x.n} (${x.k}, min ${x.min} pt)`).join(", "));
    // 2. alt text
    const noAlt = slides.filter(s => s.noAlt);
    add("A2", "Images have alternative text", noAlt.length ? "Partly" : "Fits", noAlt.length ? `${noAlt.reduce((a, s) => a + s.noAlt, 0)} image(s) without a description.` : "Every image has a description.", noAlt.map(s => `slide ${s.n}`).join(", "));
    // 3. colour alone / contrast: manual
    add("A3", "Never colour alone; high contrast", "To check", "Cannot be measured from the file. Check on a projector: every colour-coded meaning also has a label, a shape or a pattern.");
    // 4. acronyms
    const ac = {}; slides.forEach(s => acronyms(s.text + "\n" + s.notes).forEach(a => { (ac[a] = ac[a] || new Set()).add(s.n); }));
    const acList = Object.keys(ac).sort();
    add("D1", "No acronyms, on the slides or in the speech", acList.length ? "Does not fit" : "Fits", acList.length ? `${acList.length} acronym(s): ${acList.join(", ")}.` : "None found.", acList.map(a => `${a}: ${[...ac[a]].join(",")}`).join("; "));
    // 5. outline slide
    const outl = slides.filter(s => OUTLINE.test(s.title.trim()));
    add("S1", "No outline slide", outl.length ? "Does not fit" : "Fits", outl.length ? "An outline or agenda slide was found." : "None found.", outl.map(s => `slide ${s.n}`).join(", "));
    // 6. placeholders
    const ph = []; slides.forEach(s => { for (const m of (s.text + "\n" + s.notes).match(/\[[^\]\n]{1,40}\]|_{3,}|\bTBD\b|\bTODO\b|\bXX+\b|lorem ipsum/gi) || []) if (!/\d:\d\d/.test(m) && !/^\[(pause|click|let |keywords?|sources?|note|cue|transition)\b/i.test(m)) ph.push({ n: s.n, m }); });
    add("P1", "No placeholders left", ph.length ? "Does not fit" : "Fits", ph.length ? `${ph.length} placeholder(s): ${[...new Set(ph.map(x => x.m))].join(", ")}.` : "None found.", [...new Set(ph.map(x => `slide ${x.n}`))].join(", "));
    // 7. links
    const lk = []; slides.forEach(s => [...new Set(s.links.map(l => l.replace(/^https?:\/\//, '').replace(/\/$/, '')))].forEach(l => lk.push({ n: s.n, l, personal: PERSONAL_HOSTS.test(l) })));
    const pers = lk.filter(x => x.personal);
    add("L1", "Links work on the day, and are hosted by the team or the department", pers.length ? "Partly" : (lk.length ? "To check" : "Fits"),
      pers.length ? `${pers.length} link(s) point to personal or third-party hosts, with no guarantee of access on the day: ${[...new Set(pers.map(x => x.l))].join(", ")}.` : (lk.length ? `${lk.length} link(s): open each one from the presentation room's network.` : "No links."),
      [...new Set(lk.map(x => `slide ${x.n}`))].join(", "));
    // 8. timing from notes
    const marks = []; slides.forEach(s => { const re = /(\d{1,2}):(\d{2})\s*[–-]\s*(\d{1,2}):(\d{2})/g; let m; while ((m = re.exec(s.notes))) marks.push({ n: s.n, a: +m[1] * 60 + +m[2], b: +m[3] * 60 + +m[4] }); });
    const mainEnd = marks.length ? Math.max(...marks.filter(x => x.b <= 900).map(x => x.b)) : null;
    const fmt = t => `${Math.floor(t / 60)} min ${String(t % 60).padStart(2, "0")} s`;
    add("T1", "Use the full 10 minutes", mainEnd === null ? "To check" : (mainEnd > 600 ? "Does not fit" : (mainEnd >= 540 ? "Fits" : "Partly")),
      mainEnd === null ? "No timing marks (m:ss – m:ss) found in the speaker notes. Time a rehearsal." : `Scripted length: ${fmt(mainEnd)}` + (mainEnd > 600 ? ", over 10 minutes." : mainEnd >= 540 ? "." : `; about ${fmt(600 - mainEnd)} unused.`));
    // 9. first 60 seconds: show content for self-check
    const first = marks.filter(x => x.a < 60).map(x => x.n);
    const firstSlides = first.length ? [...new Set(first)] : [1, 2];
    add("T2", "The first 60 seconds say what the problem is and why it matters now", "To check",
      "Read what is said in the first minute (below) and confirm both questions are answered, and the risk of the status quo is stated.", firstSlides.map(n => `slide ${n}`).join(", "));
    // 10. text density
    const dense = slides.filter(s => words(s.text) > 70);
    add("S2", "Simple, uncluttered slides; detail in annexes", dense.length ? "Partly" : "Fits", dense.length ? `${dense.length} slide(s) carry more than 70 words.` : "No slide over 70 words.", dense.map(s => `slide ${s.n} (${words(s.text)} words)`).join(", "));
    // 11. structure keywords (presence only)
    const all = slides.map(s => (s.title + " " + s.text + " " + s.notes).toLowerCase()).join(" ");
    const parts = [["risk", /\brisk|obstacle|challenge|mitigat/], ["user", /\buser|duty officer|public servant|client|canadians/], ["consulted", /consult|interview|met with|stakeholder/], ["next steps", /next step|pilot|timeline|milestone|roadmap/], ["return on investment", /\broi\b|return on investment|cost|saving|\$/]];
    const missing = parts.filter(p => !p[1].test(all)).map(p => p[0]);
    add("S3", "Structure: problem, solution, how, user, future, next steps, risks", missing.length ? "Partly" : "To check",
      missing.length ? `No mention found of: ${missing.join(", ")}. Judges ask for obstacles and risks to be named, not hidden.` : "Every part is mentioned somewhere. Check that each one has its own clear moment.");
    return { slides: slides.map(s => ({ n: s.n, title: s.title, words: words(s.text), notesWords: words(s.notes), notes: s.notes })), results: R, firstSlides };
  }

  root.PitchCheck = { analyse, check, acronyms };
  if (typeof module !== "undefined") module.exports = root.PitchCheck;
})(typeof window !== "undefined" ? window : globalThis);
