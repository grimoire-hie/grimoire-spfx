/**
 * SentimentAnalyzer
 * Lightweight keyword/pattern-based sentiment detection for auto-triggering
 * facial expressions on the Grimm particle avatar.
 *
 * NOT a full NLP pipeline — just fast pattern matching tuned for
 * M365 assistant conversations (search, documents, sites, people).
 */

import { Expression } from './ExpressionEngine';
import { SENTIMENT_CLASSIFIER_PROMPT } from '../../config/promptCatalog';
import { getRuntimeTuningConfig } from '../config/RuntimeTuningConfig';
import type { NanoService } from '../nano/NanoService';

// ─── Pattern Rules ──────────────────────────────────────────────

interface ISentimentRule {
  /** Expression to trigger */
  expression: Expression;
  /** Regex patterns (tested against lowercased text) */
  patterns: RegExp[];
  /** Priority — higher wins when multiple rules match */
  priority: number;
  /** Duration in ms before reverting to idle (0 = don't auto-revert) */
  revertMs: number;
}

/**
 * Rules for USER messages (what the user says).
 * These reflect the user's emotional state / intent.
 */
const USER_RULES: ISentimentRule[] = [
  // ─── Happy / Gratitude ─────────────────────────────────────
  {
    expression: 'happy',
    patterns: [
      /\b(thanks?|thank\s*you|thx|appreciate|awesome|great|perfect|love\s*it|wonderful|amazing|excellent|fantastic|brilliant|nice|cool|sweet)\b/,
      /\b(well\s*done|good\s*job|that'?s?\s*(great|perfect|awesome|exactly))\b/,
      /\b(yes!|yay|woo|hooray)\b/i,
      // fr
      /\b(merci|g[eé]nial|parfait|super|formidable|bravo|magnifique|excellent)\b/i,
      // it
      /\b(grazie|fantastico|perfetto|ottimo|bravissimo|meraviglioso|eccellente|stupendo)\b/i,
      // de
      /\b(danke|toll|perfekt|super|wunderbar|ausgezeichnet|klasse|prima|hervorragend)\b/i,
      // es
      /\b(gracias|genial|perfecto|estupendo|maravilloso|excelente|fant[aá]stico|incre[ií]ble)\b/i
    ],
    priority: 10,
    revertMs: 2000
  },

  // ─── Surprised / Impressed ─────────────────────────────────
  {
    expression: 'surprised',
    patterns: [
      /\b(wow|whoa|oh\s*my|no\s*way|really\??|seriously\??|incredible|unbelievable)\b/,
      /\b(didn'?t\s*(know|expect)|i\s*had\s*no\s*idea)\b/,
      /!{2,}/,  // Multiple exclamation marks
      // fr
      /\b(incroyable|impressionnant|s[eé]rieusement|pas possible|oh l[aà] l[aà])\b/i,
      // it
      /\b(incredibile|impressionante|davvero|sul serio|non ci credo|mamma mia)\b/i,
      // de
      /\b(unglaublich|beeindruckend|echt|im ernst|wahnsinn|donnerwetter)\b/i,
      // es
      /\b(incre[ií]ble|impresionante|en serio|no puede ser|madre m[ií]a)\b/i
    ],
    priority: 8,
    revertMs: 1500
  },

  // ─── Confused / Uncertain ──────────────────────────────────
  {
    expression: 'confused',
    patterns: [
      /\b(confused|don'?t\s*understand|what\s*(do\s*you\s*mean|is\s*that)|huh\??|unclear)\b/,
      /\b(i'?m?\s*(not\s*sure|unsure|lost)|can\s*you\s*(explain|clarify|repeat))\b/,
      /\b(what\s*does\s*that\s*mean|how\s*does\s*that\s*work|i\s*don'?t\s*(get|know))\b/,
      /^\s*\?\s*$/,  // Just a question mark
      // fr
      /\b(confus|je ne comprends pas|pas clair|peux-tu expliquer|qu'est-ce que)\b/i,
      // it
      /\b(confus[oa]|non capisco|non [eè] chiaro|puoi spiegare|cosa significa)\b/i,
      // de
      /\b(verwirrt|verstehe nicht|unklar|kannst du erkl[aä]ren|was meinst du)\b/i,
      // es
      /\b(confundid[oa]|no entiendo|no est[aá] claro|puedes explicar|qu[eé] significa)\b/i
    ],
    priority: 9,
    revertMs: 2500
  },

  // ─── Thinking / Deliberating ───────────────────────────────
  {
    expression: 'thinking',
    patterns: [
      /\b(hmm+|let\s*me\s*think|good\s*question|interesting)\b/,
      /\b(i'?m?\s*(thinking|wondering|considering)|maybe|perhaps|not\s*sure\s*(yet|about))\b/,
      /\b(what\s*(if|about)|how\s*about|should\s*(i|we))\b/,
      // fr
      /\b(voyons|laisse-moi r[eé]fl[eé]chir|bonne question|int[eé]ressant|peut-[eê]tre)\b/i,
      // it
      /\b(vediamo|fammi pensare|buona domanda|interessante|forse)\b/i,
      // de
      /\b(mal sehen|lass mich [uü]berlegen|gute frage|interessant|vielleicht)\b/i,
      // es
      /\b(veamos|d[eé]jame pensar|buena pregunta|interesante|quiz[aá]s|tal vez)\b/i
    ],
    priority: 5,
    revertMs: 2000
  }
];

/**
 * Rules for ASSISTANT messages (what Grimm says).
 * These reflect Grimm's emotional state during the response.
 */
const ASSISTANT_RULES: ISentimentRule[] = [
  // ─── Happy / Success ───────────────────────────────────────
  {
    expression: 'happy',
    patterns: [
      /\b(done|completed|found|retrieved|showing|loaded|connected|here\s*(are|is))\b/,
      /\b(great\s*choice|excellent|perfect|wonderful)\b/,
      /\b(results?\s*(are|is)\s*(ready|here|showing))\b/i,
      // fr
      /\b(termin[eé]|trouv[eé]|charg[eé]|connect[eé]|voici)\b/i,
      // it
      /\b(fatto|trovato|caricato|connesso|ecco)\b/i,
      // de
      /\b(fertig|gefunden|geladen|verbunden|hier (?:sind|ist))\b/i,
      // es
      /\b(listo|encontrado|cargado|conectado|aqu[ií] (?:est[aá]n|est[aá]))\b/i
    ],
    priority: 6,
    revertMs: 1500
  },

  // ─── Thinking / Processing ─────────────────────────────────
  {
    expression: 'thinking',
    patterns: [
      /\b(let\s*me|i'?ll\s*(check|look|think|figure|find))\b/,
      /(?:^|[.!]\s*)(analyzing|processing|checking|searching|looking\s*(into|at|for)|browsing|querying|connecting)\b/,
      /\b(i'?m\s+|i\s+am\s+|currently\s+)(analyzing|processing|checking|searching|looking\s*(into|at|for)|browsing|querying|connecting)\b/,
      /\b(one\s*moment|give\s*me\s*a\s*second|hold\s*on)\b/,
      // fr
      /\b(laissez-moi|je vais v[eé]rifier|un instant|je cherche|je regarde)\b/i,
      // it
      /\b(lasciami|controllo|un momento|sto cercando|sto guardando)\b/i,
      // de
      /\b(lass mich|ich pr[uü]fe|einen moment|ich suche|ich schaue)\b/i,
      // es
      /\b(d[eé]jame|voy a verificar|un momento|estoy buscando|estoy mirando)\b/i
    ],
    priority: 7,
    revertMs: 2000
  },

  // ─── Confused / Needs clarification ────────────────────────
  {
    expression: 'confused',
    patterns: [
      /\b(could\s*you\s*(clarify|explain|tell\s*me\s*more))\b/,
      /\b(i'?m?\s*not\s*(sure|certain)|what\s*(exactly|specifically))\b/,
      /\b(can\s*you\s*be\s*more\s*specific|do\s*you\s*mean)\b/,
      // fr
      /\b(pourriez-vous pr[eé]ciser|je ne suis pas s[uû]r|que voulez-vous dire)\b/i,
      // it
      /\b(puoi chiarire|non sono sicuro|cosa intendi)\b/i,
      // de
      /\b(k[oö]nnten sie pr[aä]zisieren|ich bin nicht sicher|was meinen sie)\b/i,
      // es
      /\b(puede aclarar|no estoy seguro|qu[eé] quiere decir)\b/i
    ],
    priority: 8,
    revertMs: 2000
  },

  // ─── Surprised / Unexpected ────────────────────────────────
  {
    expression: 'surprised',
    patterns: [
      /\b(interesting|that'?s?\s*(an?\s+)?(unusual|unexpected|creative|ambitious))\b/,
      /\b(i\s*haven'?t\s*seen\s*that\s*before|that'?s?\s*new)\b/,
      // fr
      /\b(int[eé]ressant|inattendu|cr[eé]atif|ambitieux)\b/i,
      // it
      /\b(interessante|inaspettato|creativo|ambizioso)\b/i,
      // de
      /\b(interessant|unerwartet|kreativ|ambitioniert)\b/i,
      // es
      /\b(interesante|inesperado|creativo|ambicioso)\b/i
    ],
    priority: 5,
    revertMs: 1500
  }
];

// ─── Analyzer ────────────────────────────────────────────────────

export interface ISentimentResult {
  /** Detected expression, or undefined if no strong signal */
  expression: Expression | undefined;
  /** How long to hold the expression before reverting to idle (ms) */
  revertMs: number;
  /** Confidence indicator (matched pattern count) */
  matchCount: number;
}

/**
 * Analyze a transcript message and return a suggested expression.
 *
 * @param text  - The transcript text
 * @param role  - Who said it: 'user' or 'assistant'
 * @returns Suggested expression with revert timing, or undefined expression if neutral
 */
export function analyzeSentiment(
  text: string,
  role: 'user' | 'assistant' | 'system'
): ISentimentResult {
  // System messages don't trigger expressions
  if (role === 'system') {
    return { expression: undefined, revertMs: 0, matchCount: 0 };
  }

  const lower = text.toLowerCase();
  const rules = role === 'user' ? USER_RULES : ASSISTANT_RULES;

  let bestExpression: Expression | undefined;
  let bestPriority = -1;
  let bestRevertMs = 0;
  let totalMatches = 0;

  for (const rule of rules) {
    let matched = false;
    for (const pattern of rule.patterns) {
      if (pattern.test(lower)) {
        matched = true;
        totalMatches++;
        break; // One match per rule is enough
      }
    }

    if (matched && rule.priority > bestPriority) {
      bestExpression = rule.expression;
      bestPriority = rule.priority;
      bestRevertMs = rule.revertMs;
    }
  }

  return {
    expression: bestExpression,
    revertMs: bestRevertMs,
    matchCount: totalMatches
  };
}

// ─── LLM-Enhanced Sentiment (Nano) ──────────────────────────────

const VALID_EXPRESSIONS: ReadonlyArray<string> = ['happy', 'surprised', 'confused', 'thinking', 'idle'];

/**
 * Async Nano-enhanced sentiment analysis.
 * Tries regex first (instant), then calls Nano for ambiguous/unmatched text.
 * Returns the regex result immediately if it matches; otherwise awaits Nano.
 */
export async function analyzeSentimentAsync(
  text: string,
  role: 'user' | 'assistant',
  nanoService: NanoService | undefined
): Promise<ISentimentResult> {
  // 1. Try regex first (instant)
  const regexResult = analyzeSentiment(text, role);
  if (regexResult.expression) return regexResult;

  // 2. If no regex match and Nano available, try LLM classification
  if (!nanoService) return regexResult;

  const tuning = getRuntimeTuningConfig().nano;
  const response = await nanoService.classify(
    SENTIMENT_CLASSIFIER_PROMPT,
    `Role: ${role}\nText: "${text}"`,
    tuning.sentimentTimeoutMs
  );

  // 3. Parse response: "happy 2000" or "confused 2500" or "none"
  if (response && response !== 'none') {
    const parts = response.split(' ');
    const expr = parts[0];
    if (VALID_EXPRESSIONS.indexOf(expr) !== -1) {
      return {
        expression: expr as Expression,
        revertMs: parseInt(parts[1], 10) || 2000,
        matchCount: 1
      };
    }
  }

  return regexResult;
}

/**
 * Quick check: does the text contain a question the user is asking Grimm?
 * Used to set 'listening' expression proactively.
 */
export function isQuestion(text: string): boolean {
  const t = text.trim();
  return /\?\s*$/.test(t)
    || /^(what|how|why|when|where|who|which|can|could|would|should|do|does|is|are|will)\b/i.test(t)
    // fr
    || /^(qu['e]|comment|pourquoi|quand|o[uù]|qui|quel|est-ce)\b/i.test(t)
    // it
    || /^(che|come|perch[eé]|quando|dove|chi|quale)\b/i.test(t)
    // de
    || /^(was|wie|warum|wann|wo|wer|welche[rs]?|kann|k[oö]nnte)\b/i.test(t)
    // es
    || /^(qu[eé]|c[oó]mo|por qu[eé]|cu[aá]ndo|d[oó]nde|qui[eé]n|cu[aá]l|puede)\b/i.test(t);
}
