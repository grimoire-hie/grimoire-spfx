export type HiePromptMessageKind =
  | 'visual'
  | 'visual-current-state'
  | 'task'
  | 'interaction'
  | 'flow'
  | 'tool-completion'
  | 'tool-error';

export interface IHiePromptMessage {
  kind: HiePromptMessageKind;
  body: string;
}

const PROMPT_PREFIX_BY_KIND: Record<HiePromptMessageKind, string> = {
  visual: '[Visual context: ',
  'visual-current-state': '[Visual context (current state): ',
  task: '[Task context: ',
  interaction: '[User interaction:',
  flow: '[Flow update: ',
  'tool-completion': '[Tool completed: ',
  'tool-error': '[Tool error: '
};

export function formatHiePromptMessage(message: IHiePromptMessage): string {
  if (message.kind === 'interaction') {
    return [
      PROMPT_PREFIX_BY_KIND.interaction,
      message.body.trim(),
      ']'
    ].join('\n');
  }

  return `${PROMPT_PREFIX_BY_KIND[message.kind]}${message.body.trim()}]`;
}

export function parseHiePromptMessage(text: string): IHiePromptMessage | undefined {
  const normalizedText = text.trim();
  if (!normalizedText.startsWith('[') || !normalizedText.endsWith(']')) {
    return undefined;
  }

  if (normalizedText.startsWith(PROMPT_PREFIX_BY_KIND.interaction)) {
    const inner = normalizedText
      .slice(PROMPT_PREFIX_BY_KIND.interaction.length, -1)
      .trim();
    return inner ? { kind: 'interaction', body: inner } : undefined;
  }

  const nonInteractionKinds = Object.keys(PROMPT_PREFIX_BY_KIND)
    .filter((kind): kind is Exclude<HiePromptMessageKind, 'interaction'> => kind !== 'interaction');

  for (let i = 0; i < nonInteractionKinds.length; i++) {
    const kind = nonInteractionKinds[i];
    const prefix = PROMPT_PREFIX_BY_KIND[kind];
    if (normalizedText.startsWith(prefix)) {
      const inner = normalizedText.slice(prefix.length, -1).trim();
      return inner ? { kind, body: inner } : undefined;
    }
  }

  return undefined;
}

export function isHiePromptMessage(text: string): boolean {
  return parseHiePromptMessage(text) !== undefined;
}

export function isHieToolErrorPromptMessage(text: string): boolean {
  return parseHiePromptMessage(text)?.kind === 'tool-error';
}
