jest.mock('@fluentui/react/lib/Dropdown', () => {
  return {
    Dropdown: () => null
  };
});

jest.mock('@fluentui/react/lib/Button', () => {
  return {
    IconButton: () => null
  };
});

jest.mock('@fluentui/react/lib/Toggle', () => {
  return {
    Toggle: () => null
  };
});

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act } from 'react-dom/test-utils';

import { AvatarSettingsPanel } from './AvatarSettingsPanel';
import { DEFAULT_SP_THEME_COLORS, useGrimoireStore } from '../../store/useGrimoireStore';

describe('AvatarSettingsPanel visage persona helper', () => {
  let container: HTMLDivElement;

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
    useGrimoireStore.setState({
      avatarEnabled: true,
      voiceId: 'alloy',
      publicWebSearchEnabled: false,
      publicWebSearchCapability: 'unknown',
      publicWebSearchCapabilityDetail: undefined,
      copilotWebGroundingEnabled: false,
      searchRecapMode: 'auto',
      personality: 'normal',
      visage: 'cat',
      spThemeColors: DEFAULT_SP_THEME_COLORS
    });
  });

  afterEach(() => {
    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    container.remove();
    jest.clearAllMocks();
  });

  it('renders and updates the selected avatar persona helper text', async () => {
    await act(async () => {
      ReactDom.render(
        <AvatarSettingsPanel
          isOpen={true}
          onDismiss={() => undefined}
        />,
        container
      );
      await Promise.resolve();
    });

    expect(container.textContent).toContain(
      'Majestic Cat - Elegant, polite, slightly distant overseer of Microsoft 365 services.'
    );
    expect(container.textContent).toContain(
      'Voice changes apply the next time you connect voice.'
    );

    await act(async () => {
      useGrimoireStore.getState().setVisage('squirrel');
      await Promise.resolve();
    });

    expect(container.textContent).toContain(
      'Squirrel - Nimble, funny, lightly mischievous helper that stays competent.'
    );
    expect(container.textContent).not.toContain(
      'Majestic Cat - Elegant, polite, slightly distant overseer of Microsoft 365 services.'
    );
  });
});
