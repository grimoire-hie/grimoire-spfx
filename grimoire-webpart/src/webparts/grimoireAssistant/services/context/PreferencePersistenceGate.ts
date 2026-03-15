export interface IPreferencePersistenceDecision {
  nextBaseline: string;
  shouldPersist: boolean;
}

export function getPreferencePersistenceDecision(
  baselinePayload: string | undefined,
  currentPayload: string
): IPreferencePersistenceDecision {
  if (typeof baselinePayload === 'undefined') {
    return {
      nextBaseline: currentPayload,
      shouldPersist: false
    };
  }

  return {
    nextBaseline: baselinePayload,
    shouldPersist: baselinePayload !== currentPayload
  };
}
