import { emitBlockInteraction, resolveInteractionSchemaId } from './interactionSchemas';
import { hybridInteractionEngine } from '../../services/hie/HybridInteractionEngine';

describe('interactionSchemas', () => {
  it('resolves schema IDs for core actions', () => {
    expect(resolveInteractionSchemaId('selection-list', 'select')).toBe('selection.select');
    expect(resolveInteractionSchemaId('form', 'submit-form')).toBe('form.submit');
    expect(resolveInteractionSchemaId('search-results', 'click-result')).toBe('search-results.click-result');
    expect(resolveInteractionSchemaId('document-library', 'click-file')).toBe('document-library.click-file');
    expect(resolveInteractionSchemaId('permissions-view', 'click-permission')).toBe('permissions-view.click-permission');
    expect(resolveInteractionSchemaId('user-card', 'click-user')).toBe('user-card.click-user');
    expect(resolveInteractionSchemaId('search-results', 'look')).toBe('hover.look');
  });

  it('emits normalized interactions through HIE', () => {
    const spy = jest.spyOn(hybridInteractionEngine, 'onBlockInteraction').mockImplementation(() => { /* no-op */ });

    const emitted = emitBlockInteraction({
      blockId: 'block-1',
      blockType: 'selection-list',
      action: 'select',
      payload: { label: 'A' },
      schemaId: 'selection.select'
    });

    expect(emitted).toBe(true);
    expect(spy).toHaveBeenCalledTimes(1);
    expect(spy.mock.calls[0][0].schemaId).toBe('selection.select');
    expect(spy.mock.calls[0][0].eventName).toBe('block.interaction.select');
    expect(spy.mock.calls[0][0].exposurePolicy).toMatchObject({ mode: 'response-triggering' });

    spy.mockRestore();
  });
});
