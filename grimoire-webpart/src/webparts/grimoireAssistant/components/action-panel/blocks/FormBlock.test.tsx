const peoplePickerMock: jest.Mock = jest.fn(() => null);

jest.mock('@pnp/spfx-controls-react/lib/TeamPicker', () => ({
  TeamPicker: () => null
}));

jest.mock('@pnp/spfx-controls-react/lib/TeamChannelPicker', () => ({
  TeamChannelPicker: () => null
}));

jest.mock('@pnp/spfx-controls-react/lib/PeoplePicker', () => {
  return {
    PeoplePicker: peoplePickerMock,
    PrincipalType: {
      User: 1
    }
  };
});

jest.mock('../../../services/pnp/pnpContext', () => ({
  getContext: jest.fn(() => undefined)
}));

jest.mock('../interactionSchemas', () => ({
  emitBlockInteraction: jest.fn()
}));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act } from 'react-dom/test-utils';

import type { IFormData } from '../../../models/IBlock';
import { FormBlock } from './FormBlock';
import { getContext } from '../../../services/pnp/pnpContext';

describe('FormBlock attachments', () => {
  let container: HTMLDivElement;
  let openSpy: jest.SpyInstance;

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
    peoplePickerMock.mockReset().mockImplementation(() => null);
    openSpy = jest.spyOn(window, 'open').mockImplementation(() => null);
  });

  afterEach(() => {
    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    container.remove();
    openSpy.mockRestore();
    jest.clearAllMocks();
  });

  it('allows removing prepared attachments before submit', async () => {
    const onSubmit = jest.fn().mockResolvedValue({ success: true, message: 'Sent.' });
    const onUpdateBlock = jest.fn();
    const data: IFormData = {
      kind: 'form',
      preset: 'email-compose',
      description: 'Share recap',
      fields: [],
      submissionTarget: {
        toolName: 'SendEmailWithAttachments',
        serverId: 'mcp_MailTools',
        staticArgs: {
          attachmentUris: [
            'https://tenant.sharepoint.com/sites/dev/SPFx_ja.pdf',
            'https://tenant.sharepoint.com/sites/dev/SPFx_de.pdf'
          ]
        }
      },
      status: 'editing'
    };

    await act(async () => {
      ReactDom.render(
        <FormBlock
          data={data}
          blockId="block-form"
          onSubmit={onSubmit}
          onUpdateBlock={onUpdateBlock}
        />,
        container
      );
      await Promise.resolve();
    });

    const openButton = Array.from(container.querySelectorAll('button')).find((button) => button.textContent === 'SPFx_ja.pdf') as HTMLButtonElement | undefined;
    expect(openButton).toBeDefined();

    await act(async () => {
      openButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      await Promise.resolve();
    });

    expect(openSpy).toHaveBeenCalledWith(
      'https://tenant.sharepoint.com/sites/dev/SPFx_ja.pdf',
      '_blank',
      'noopener,noreferrer'
    );

    const removeButton = container.querySelector('button[aria-label="Remove attachment SPFx_ja.pdf"]') as HTMLButtonElement | null;
    expect(removeButton).not.toBeNull();

    await act(async () => {
      removeButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      await Promise.resolve();
    });

    expect(container.textContent).not.toContain('SPFx_ja.pdf');
    expect(container.textContent).toContain('SPFx_de.pdf');

    const submitButton = Array.from(container.querySelectorAll('button')).find((button) => button.textContent === 'Submit') as HTMLButtonElement | undefined;
    expect(submitButton).toBeDefined();

    await act(async () => {
      submitButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      await Promise.resolve();
    });

    expect(onSubmit).toHaveBeenCalledTimes(1);
    expect((onSubmit.mock.calls[0]?.[0] as IFormData).submissionTarget.staticArgs).toEqual({
      attachmentUris: ['https://tenant.sharepoint.com/sites/dev/SPFx_de.pdf']
    });
  });
});

describe('FormBlock Teams people picker', () => {
  let container: HTMLDivElement;

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
    peoplePickerMock.mockReset().mockImplementation(({ onChange }: { onChange?: (items: Array<{ text?: string; secondaryText?: string }>) => void }) => (
      <button type="button" onClick={() => onChange?.([{ text: 'Test User', secondaryText: 'test.user@example.com' }])}>
        Mock PeoplePicker
      </button>
    ));
    (getContext as jest.Mock).mockReturnValue({
      pageContext: {
        web: {
          absoluteUrl: 'https://tenant.sharepoint.com/sites/dev'
        }
      },
      msGraphClientFactory: {},
      spHttpClient: {}
    });
  });

  afterEach(() => {
    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    container.remove();
    jest.clearAllMocks();
  });

  it('submits Teams chat recipients selected from the people picker', async () => {
    const onSubmit = jest.fn().mockResolvedValue({ success: true, message: 'Shared.' });
    const data: IFormData = {
      kind: 'form',
      preset: 'share-teams-chat',
      description: 'Share to Teams chat',
      fields: [
        { key: 'recipients', label: 'People', type: 'people-picker', required: true, placeholder: 'Search people' },
        { key: 'content', label: 'Message', type: 'textarea', required: true, defaultValue: 'Review this in Teams.' }
      ],
      submissionTarget: {
        toolName: 'share_teams_chat',
        serverId: 'internal-share',
        staticArgs: {}
      },
      status: 'editing'
    };

    await act(async () => {
      ReactDom.render(
        <FormBlock
          data={data}
          onSubmit={onSubmit}
        />,
        container
      );
      await Promise.resolve();
    });

    const pickerButton = Array.from(container.querySelectorAll('button')).find((button) => button.textContent === 'Mock PeoplePicker') as HTMLButtonElement | undefined;
    expect(pickerButton).toBeDefined();

    await act(async () => {
      pickerButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      await Promise.resolve();
    });

    const submitButton = Array.from(container.querySelectorAll('button')).find((button) => button.textContent === 'Submit') as HTMLButtonElement | undefined;
    expect(submitButton).toBeDefined();

    await act(async () => {
      submitButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      await Promise.resolve();
    });

    expect(onSubmit).toHaveBeenCalledTimes(1);
    expect(onSubmit.mock.calls[0]?.[2]).toEqual({
      recipients: ['test.user@example.com']
    });
  });
});
