import { spawn } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { pathToFileURL } from "node:url";

const chrome = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";
const gamePath = resolve("outputs/island-republic-game/index.html");
const gameUrl = pathToFileURL(gamePath).href;
const profile = mkdtempSync(join(tmpdir(), "ir-chrome-test-"));
const port = 9323;
const errors = [];
const results = [];
let browser;

const sleep = ms => new Promise(resolvePromise => setTimeout(resolvePromise, ms));
const assert = (condition, message) => {
  if (!condition) throw new Error(message);
  results.push(`PASS: ${message}`);
};

class CDP {
  constructor(url) {
    this.ws = new WebSocket(url);
    this.id = 0;
    this.pending = new Map();
    this.ready = new Promise((resolveReady, rejectReady) => {
      this.ws.onopen = resolveReady;
      this.ws.onerror = rejectReady;
    });
    this.ws.onmessage = event => {
      const message = JSON.parse(event.data);
      if (message.id && this.pending.has(message.id)) {
        const { resolve: resolvePending, reject } = this.pending.get(message.id);
        this.pending.delete(message.id);
        if (message.error) reject(new Error(message.error.message));
        else resolvePending(message.result);
      } else if (message.method === "Runtime.exceptionThrown") {
        errors.push(message.params.exceptionDetails.text);
      }
    };
  }
  async send(method, params = {}) {
    await this.ready;
    const id = ++this.id;
    return new Promise((resolvePending, reject) => {
      this.pending.set(id, { resolve: resolvePending, reject });
      this.ws.send(JSON.stringify({ id, method, params }));
    });
  }
  close() { this.ws.close(); }
}

async function findPageTarget() {
  for (let attempt = 0; attempt < 80; attempt++) {
    try {
      const targets = await fetch(`http://127.0.0.1:${port}/json`).then(r => r.json());
      const target = targets.find(t => t.type === "page" && t.url.startsWith("file:"));
      if (target) return target;
    } catch {}
    await sleep(100);
  }
  throw new Error("Chrome DevTools target did not become available");
}

async function main() {
  browser = spawn(chrome, [
    "--headless=new",
    "--disable-gpu",
    "--no-first-run",
    "--no-default-browser-check",
    `--user-data-dir=${profile}`,
    `--remote-debugging-port=${port}`,
    gameUrl
  ], { stdio: "ignore", windowsHide: true });

  const target = await findPageTarget();
  const cdp = new CDP(target.webSocketDebuggerUrl);
  await cdp.send("Runtime.enable");
  await cdp.send("Page.enable");

  const evaluate = async expression => {
    const out = await cdp.send("Runtime.evaluate", { expression, returnByValue: true, awaitPromise: true });
    if (out.exceptionDetails) throw new Error(out.exceptionDetails.text);
    return out.result.value;
  };
  const waitFor = async (expression, message, attempts = 80) => {
    for (let i = 0; i < attempts; i++) {
      if (await evaluate(expression)) return;
      await sleep(50);
    }
    throw new Error(`Timed out: ${message}`);
  };
  const click = async selector => {
    const ok = await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e)return false;e.click();return true})()`);
    assert(ok, `可操作 ${selector}`);
    await sleep(30);
  };
  const text = () => evaluate("document.body.innerText");
  const state = () => evaluate("window.__IR_TEST__.getState()");

  await waitFor("document.readyState==='complete' && !!window.__IR_TEST__", "game bootstrap");
  assert((await text()).includes("現成法案"), "首頁顯示現成法案流程");
  assert(!/(拼裝主要法案|強制總預算|利益 1|舉債取得)/.test(await text()), "首頁沒有舊拼裝或預算介面");

  await click("#startBtn");
  assert((await state()).factions.every(f => f.seats === 5), "開局四個政治網絡席次完全相同");
  const initial = await state();
  assert(["production","capacity","trust","pressure"].every(k => !(k in initial)), "國家狀態已移除生產、行政、信任與分配壓力");
  assert(initial.factions.every(f => ["exposure","leadership","integrity"].every(k => !(k in f))), "政治網絡已移除曝光、國家領導與清廉");
  assert((await evaluate("document.querySelectorAll('#metrics .metric').length")) === 3, "頂部只保留國庫、債務與危機三項國家指標");
  await evaluate("(()=>{const sequence=[0.34,0.76];Math.random=()=>sequence.length?sequence.shift():0})()");
  await click("#drawEvent");
  assert((await text()).includes("提高居住可負擔性"), "事件形成可見的社會需求");
  assert((await state()).demands.length === 1, "事件建立需求狀態");
  assert((await state()).challenges.length === 1, "奇數回合揭露一項迫近國家挑戰");
  assert((await state()).challenges[0].stage === "looming", "新國家挑戰從迫近階段開始");
  assert((await text()).includes("金融市場與房貸危機"), "事件畫面顯示新揭露的國家挑戰");
  await click("#eventNext");
  await click("#resolveRevenue");
  assert((await state()).phase === 2, "收入後進入院長選舉與提案階段");
  assert(await evaluate("document.querySelectorAll('[data-chair-vote]').length") === 4, "四個政治網絡各有一張院長選票");
  await evaluate("document.querySelectorAll('[data-chair-vote]').forEach(e=>e.value='0')");
  await click("#resolveChairElection");
  assert((await state()).readyLaw.chairId === 0, "四方投票共同選出本輪院長");
  assert((await text()).includes("藍色網絡以 4 票當選"), "提案畫面顯示院長選舉結果");
  assert(!/(拼裝主要法案|強制總預算|利益 1|舉債取得)/.test(await text()), "提案畫面沒有舊拼裝欄位");

  const proposals = ["tax", "housing", "childcare", "defense"];
  for (let id = 0; id < proposals.length; id++) {
    await evaluate(`(()=>{const s=document.querySelector('[data-law-select="${id}"]');s.value=${JSON.stringify(proposals[id])};document.querySelector('[data-submit-law="${id}"]').click()})()`);
    await sleep(30);
  }
  assert((await state()).proposals.length === 4, "四個政治網絡皆可登記現成法案");
  await click("#toReadyOrdering");
  assert(await evaluate("document.querySelectorAll('.agenda-order-row').length") === 4, "院長排案保留全部四項提案");
  assert(await evaluate("document.querySelectorAll('[data-up],[data-down]').length") === 8, "院長只有排序控制，沒有刪案控制");

  await evaluate("(()=>{const s=__IR_TEST__.getState();s.treasury=2;__IR_TEST__.setState(s)})()");
  await click("#lockReadyAgenda");
  assert((await text()).includes("類型："), "固定法案牌只用簡單類型連結需求與國家挑戰");
  assert(await evaluate("Object.values(__IR_TEST__.laws[0]).sort().length") === 5, "法案資料除內部識別碼外只含名稱、花費、效果與類型");
  const setAllVotes = async value => evaluate(`document.querySelectorAll('[data-vote]').forEach(e=>e.value=${JSON.stringify(value)})`);

  await setAllVotes("yes");
  await click("#resolveReadyVote");
  assert((await state()).treasury === 5, "先通過增稅法後國庫由 2 增至 5");
  assert((await state()).demands[0].progress === 1, "通過相符法案才推進社會需求");
  assert((await state()).challenges[0].progress === 1, "通過相符類型法案會增加一點國家挑戰應變值");
  await click("#continueReadyAgenda");

  await setAllVotes("yes");
  await click("#resolveReadyVote");
  assert((await state()).treasury === 2, "住宅支出扣除固定成本，解除金融挑戰再取得國庫獎勵");
  assert((await state()).demands[0].fulfilled === true, "相符法案正式通過後滿足需求");
  assert((await state()).challenges[0].resolved === true, "累積足夠應變值後解除迫近國家挑戰");
  assert((await state()).factions[1].prestige === 2, "完成挑戰的主要貢獻者取得額外聲望");
  await click("#continueReadyAgenda");

  assert((await text()).includes("財政不成立"), "國庫不足的一般支出案不進入表決");
  const beforeFiscal = await state();
  await click("#recordFiscalFailure");
  const afterFiscal = await state();
  assert(afterFiscal.treasury === beforeFiscal.treasury, "財政不成立法案不改變國庫");
  assert(afterFiscal.debt === beforeFiscal.debt && afterFiscal.crises === beforeFiscal.crises, "財政不成立法案不產生國家效果");
  await click("#continueReadyAgenda");

  const debtBefore = (await state()).debt;
  await setAllVotes("no");
  await click("#resolveReadyVote");
  assert((await state()).debt === debtBefore, "否決的特別舉債法不增加債務");
  await click("#continueReadyAgenda");
  assert((await state()).phase === 6, "全部提案處理後直接進入民意階段");

  await click("#saveBtn");
  await cdp.send("Page.reload", { ignoreCache: true });
  await waitFor("document.readyState==='complete' && !!window.__IR_TEST__ && !!document.querySelector('#loadBtn')", "reload");
  await click("#loadBtn");
  assert((await state()).phase === 6, "重新整理並讀取後可繼續原階段");
  assert((await state()).challenges[0].resolved === true, "存檔重載保留國家挑戰狀態");

  await evaluate(`(()=>{const s=__IR_TEST__.getState();s.challenges.push(
    {uid:'CTEST1',challengeId:'typhoon',stage:'looming',countdown:1,progress:0,contributors:{},revealedRound:1,resolved:false,activeTurns:0,disasterTurns:0,scarApplied:false},
    {uid:'CTEST2',challengeId:'drought',stage:'looming',countdown:1,progress:0,contributors:{},revealedRound:1,resolved:false,activeTurns:0,disasterTurns:0,scarApplied:false}
  );s.drawnChallengeIds.push('typhoon','drought');s.challengeReviewRound=null;__IR_TEST__.setState(s)})()`);
  assert((await text()).includes("連鎖危機"), "相關挑戰同時存在時顯示連鎖危機與提高門檻");

  await click("#resolveOpinion");
  assert((await state()).demands[0].resolved === true, "民意階段結算已滿足需求");
  assert((await state()).phase === 7, "民意後進入選舉階段，流程不再包含監督階段");
  await click("#skipElection");
  let challengeState = await state();
  assert(challengeState.challenges.filter(c => ["typhoon","drought"].includes(c.challengeId)).every(c => c.stage === "active"), "未處理的迫近挑戰在回合檢查時爆發");
  assert(challengeState.challengeReviewSummary.some(x => x.includes("連鎖危機") && x.includes("國庫")), "回合檢查會結算連鎖危機的國庫損失");

  const crisesBeforeDisaster = challengeState.crises;
  await evaluate("(()=>{const s=__IR_TEST__.getState();s.round=2;s.phase=8;s.challengeReviewRound=null;__IR_TEST__.setState(s)})()");
  challengeState = await state();
  assert(challengeState.challenges.filter(c => ["typhoon","drought"].includes(c.challengeId)).every(c => c.stage === "disaster"), "爆發挑戰逾期後升級為國難");
  assert(challengeState.crises === crisesBeforeDisaster + 2, "兩項國難各增加一點國家危機");

  await evaluate("(()=>{const s=__IR_TEST__.getState();s.round=3;s.phase=8;s.challengeReviewRound=null;__IR_TEST__.setState(s)})()");
  challengeState = await state();
  assert(challengeState.scars.length === 2, "國難持續未處理會留下永久國家傷痕");
  assert(errors.length === 0, "瀏覽器執行期間沒有未捕捉的 JavaScript 例外");
  assert(!/(拼裝主要法案|強制總預算|利益 1|舉債取得)/.test(await text()), "完整流程未出現舊拼裝或預算介面");

  const shot = await cdp.send("Page.captureScreenshot", { format: "png", captureBeyondViewport: true });
  writeFileSync(resolve("work/tmp/browser-test.png"), Buffer.from(shot.data, "base64"));
  cdp.close();
  console.log(results.join("\n"));
  console.log(`PASS: ${results.length} browser assertions; screenshot saved to work/tmp/browser-test.png`);
}

try {
  await main();
} finally {
  if (browser && !browser.killed) browser.kill();
  await sleep(100);
  rmSync(profile, { recursive: true, force: true });
}
