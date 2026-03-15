/**
 * NanoService — Thin HTTP client for GPT-5 Nano (fast backend).
 * Reuses the existing llmProxyHandler route: /api/fast/openai/deployments/...
 * All calls are fire-and-forget with AbortController timeouts.
 */

import type { IProxyConfig } from '../../store/useGrimoireStore';
import { getRuntimeTuningConfig } from '../config/RuntimeTuningConfig';
import { logService } from '../logging/LogService';

export class NanoService {
  private proxyUrl: string;
  private proxyApiKey: string;
  private deployment: string;
  private apiVersion: string;
  /** Cooldown timestamp after a non-timeout failure to avoid repeated console noise. */
  private disabledUntil: number = 0;
  private reasoningEffortSupported: boolean | undefined;

  constructor(proxyUrl: string, proxyApiKey: string, deployment: string, apiVersion: string) {
    this.proxyUrl = proxyUrl;
    this.proxyApiKey = proxyApiKey;
    this.deployment = deployment;
    this.apiVersion = apiVersion;
  }

  /**
   * Fire-and-forget classification call with timeout.
   * Returns parsed content string, or undefined on error/timeout.
   * Auto-disables after first failure to avoid repeated 404 console noise.
   */
  public async classify(
    systemPrompt: string,
    userMessage: string,
    timeoutMs: number = getRuntimeTuningConfig().nano.defaultTimeoutMs,
    maxTokens: number = 100
  ): Promise<string | undefined> {
    if (this.disabledUntil > Date.now()) return undefined;

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const tuning = getRuntimeTuningConfig().nano;
    const cooldownMs = tuning.cooldownMs;
    const cooldownSeconds = Math.round(cooldownMs / 1000);

    try {
      const url = `${this.proxyUrl}/fast/openai/deployments/${this.deployment}/chat/completions?api-version=${this.apiVersion}`;
      const executeRequest = async (includeReasoningEffort: boolean): Promise<Response> => {
        const requestBody: Record<string, unknown> = {
          messages: [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: userMessage }
          ],
          max_completion_tokens: maxTokens
        };

        if (includeReasoningEffort && tuning.reasoningEffort) {
          requestBody.reasoning_effort = tuning.reasoningEffort;
        }

        return fetch(url, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'api-key': this.proxyApiKey
          },
          body: JSON.stringify(requestBody),
          signal: controller.signal
        });
      };

      const shouldTryReasoningEffort = this.reasoningEffortSupported !== false && !!tuning.reasoningEffort;
      let response = await executeRequest(shouldTryReasoningEffort);
      let errBody = '';

      if (!response.ok) {
        errBody = await response.text().catch(() => '');
        if (shouldTryReasoningEffort && response.status === 400 && /reasoning[_ ]effort/i.test(errBody)) {
          this.reasoningEffortSupported = false;
          logService.info('llm', 'Nano classify retrying without reasoning_effort after upstream rejected it');
          response = await executeRequest(false);
          if (!response.ok) {
            errBody = await response.text().catch(() => '');
          }
        }
      } else if (shouldTryReasoningEffort) {
        this.reasoningEffortSupported = true;
      }

      if (!response.ok) {
        this.disabledUntil = Date.now() + cooldownMs;
        logService.info('llm', `Nano classify failed: HTTP ${response.status} — ${errBody.substring(0, 200)} — cooling down for ${cooldownSeconds}s`);
        return undefined;
      }

      const data = await response.json() as {
        choices?: Array<{ message?: { content?: string | Array<{ text?: string; type?: string }> } }>;
      };
      const content = data.choices?.[0]?.message?.content;
      if (typeof content === 'string') return content.trim();
      if (Array.isArray(content)) {
        return content
          .map((part) => typeof part?.text === 'string' ? part.text : '')
          .join('')
          .trim();
      }
      return undefined;
    } catch (err) {
      if ((err as Error).name === 'AbortError') {
        logService.debug('llm', `Nano classify timed out (${timeoutMs}ms)`);
      } else {
        this.disabledUntil = Date.now() + cooldownMs;
        logService.debug('llm', `Nano classify error: ${(err as Error).message} — cooling down for ${cooldownSeconds}s`);
      }
      return undefined;
    } finally {
      clearTimeout(timer);
    }
  }
}

// ─── Singleton Factory ──────────────────────────────────────────

let cachedInstance: NanoService | undefined;

/**
 * Get or create a NanoService instance.
 * Returns undefined if proxy config is missing.
 */
export function getNanoService(proxyConfig: IProxyConfig | undefined): NanoService | undefined {
  if (!proxyConfig) return undefined;
  if (cachedInstance) return cachedInstance;

  // Derive the fast deployment name from the current deployment's prefix.
  // e.g., "atlas-x7k2-reasoning" → prefix "atlas-x7k2" → fast "atlas-x7k2-fast"
  const prefix = proxyConfig.deployment.replace(/-reasoning$|-fast$/, '');
  const fastDeployment = prefix + '-fast';

  cachedInstance = new NanoService(
    proxyConfig.proxyUrl,
    proxyConfig.proxyApiKey,
    fastDeployment,
    proxyConfig.apiVersion
  );
  return cachedInstance;
}

/**
 * Clear the cached singleton (e.g. on disconnect).
 */
export function clearNanoService(): void {
  cachedInstance = undefined;
}
