import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act } from 'react-dom/test-utils';

import { SvgAvatar } from './SvgAvatar';
import { useGrimoireStore } from '../../store/useGrimoireStore';

const MINIMAL_SVG = `
<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
  <g id="grimoire_mascot">
    <g id="brows"></g>
    <g id="left_eye"></g>
    <g id="right_eye"></g>
    <g id="mouth"></g>
    <g id="halo"></g>
  </g>
</svg>`;

describe('SvgAvatar startup and visibility handling', () => {
  let container: HTMLDivElement;
  let hidden = false;
  let nowMs = 0;
  let nextFrameId = 0;
  let originalImage: typeof Image | undefined;
  let originalGetBBox: unknown;
  let originalDateNow: () => number;
  const rafCallbacks = new Map<number, FrameRequestCallback>();

  async function flushAsyncWork(): Promise<void> {
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();
  }

  function runNextFrame(nextNowMs?: number): void {
    if (typeof nextNowMs === 'number') {
      nowMs = nextNowMs;
    }

    const nextEntry = Array.from(rafCallbacks.entries()).sort((a, b) => a[0] - b[0])[0];
    if (!nextEntry) {
      throw new Error('No queued animation frame to execute');
    }

    const [frameId, callback] = nextEntry;
    rafCallbacks.delete(frameId);
    act(() => {
      callback(nowMs);
    });
  }

  async function waitForAvatarRenderState(expectedState: ReturnType<typeof useGrimoireStore.getState>['avatarRenderState']): Promise<void> {
    for (let attempt = 0; attempt < 8; attempt++) {
      if (useGrimoireStore.getState().avatarRenderState === expectedState) {
        return;
      }

      await act(async () => {
        await flushAsyncWork();
      });
    }

    expect(useGrimoireStore.getState().avatarRenderState).toBe(expectedState);
  }

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
    hidden = false;
    nowMs = 0;
    nextFrameId = 0;
    rafCallbacks.clear();

    useGrimoireStore.setState({
      assistantPlaybackState: 'idle',
      avatarRenderState: 'placeholder',
      avatarActionCue: undefined,
      userContext: undefined
    });

    Object.defineProperty(document, 'hidden', {
      configurable: true,
      get: () => hidden
    });

    Object.defineProperty(globalThis, 'fetch', {
      configurable: true,
      writable: true,
      value: jest.fn().mockResolvedValue({
        ok: true,
        text: async () => MINIMAL_SVG
      })
    });

    originalImage = globalThis.Image;
    class MockImage {
      public complete = false;
      public onload: ((event: Event) => void) | undefined;
      public onerror: ((event: Event | string) => void) | undefined;
      private _src = '';

      public get src(): string {
        return this._src;
      }

      public set src(value: string) {
        this._src = value;
        this.complete = true;
        Promise.resolve().then(() => {
          this.onload?.(new Event('load'));
        }).catch(() => undefined);
      }
    }

    Object.defineProperty(globalThis, 'Image', {
      configurable: true,
      writable: true,
      value: MockImage
    });

    Object.defineProperty(globalThis, 'requestAnimationFrame', {
      configurable: true,
      writable: true,
      value: jest.fn().mockImplementation((callback: FrameRequestCallback) => {
        const frameId = ++nextFrameId;
        rafCallbacks.set(frameId, callback);
        return frameId;
      })
    });
    Object.defineProperty(globalThis, 'cancelAnimationFrame', {
      configurable: true,
      writable: true,
      value: jest.fn().mockImplementation((frameId: number) => {
        rafCallbacks.delete(frameId);
      })
    });

    originalDateNow = Date.now;
    Date.now = jest.fn(() => nowMs);

    originalGetBBox = (SVGElement.prototype as unknown as { getBBox?: unknown }).getBBox;
    (SVGElement.prototype as unknown as { getBBox: () => DOMRect }).getBBox = () => ({
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      top: 0,
      right: 100,
      bottom: 100,
      left: 0,
      toJSON: () => undefined
    } as DOMRect);
  });

  afterEach(() => {
    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    container.remove();
    jest.restoreAllMocks();

    Date.now = originalDateNow;
    Object.defineProperty(globalThis, 'Image', {
      configurable: true,
      writable: true,
      value: originalImage
    });

    if (originalGetBBox) {
      (SVGElement.prototype as unknown as { getBBox: unknown }).getBBox = originalGetBBox;
    } else {
      delete (SVGElement.prototype as unknown as { getBBox?: unknown }).getBBox;
    }
  });

  it('transitions avatar render state from loading to animated-ready', async () => {
    await act(async () => {
      ReactDom.render(
        <SvgAvatar
          faceTemplate={[]}
          visage="classic"
          personality="normal"
          expression="idle"
          isActive={true}
        />,
        container
      );
      await flushAsyncWork();
    });

    await waitForAvatarRenderState('bindings-ready');

    runNextFrame(0);

    expect(useGrimoireStore.getState().avatarRenderState).toBe('animated-ready');
  });

  it('throttles idle animation work to roughly 30fps', async () => {
    await act(async () => {
      ReactDom.render(
        <SvgAvatar
          faceTemplate={[]}
          visage="classic"
          personality="normal"
          expression="idle"
          isActive={true}
        />,
        container
      );
      await flushAsyncWork();
    });

    await waitForAvatarRenderState('bindings-ready');
    runNextFrame(1000);

    const setAttributeSpy = jest.spyOn(SVGElement.prototype as unknown as { setAttribute: typeof SVGElement.prototype.setAttribute }, 'setAttribute');
    setAttributeSpy.mockClear();

    runNextFrame(1010);
    const after10Ms = setAttributeSpy.mock.calls.length;

    runNextFrame(1020);
    const after20Ms = setAttributeSpy.mock.calls.length;

    runNextFrame(1040);
    const after40Ms = setAttributeSpy.mock.calls.length;

    expect(after10Ms).toBe(0);
    expect(after20Ms).toBe(0);
    expect(after40Ms).toBeGreaterThan(0);

    ReactDom.unmountComponentAtNode(container);
  });

  it('pauses animation while hidden and resumes with a single new RAF when visible again', async () => {
    await act(async () => {
      ReactDom.render(
        <SvgAvatar
          faceTemplate={[]}
          visage="classic"
          personality="normal"
          expression="idle"
          isActive={true}
        />,
        container
      );
      await flushAsyncWork();
    });

    expect(rafCallbacks.size).toBeGreaterThan(0);
    const initialQueueSize = rafCallbacks.size;

    hidden = true;
    act(() => {
      document.dispatchEvent(new Event('visibilitychange'));
    });

    expect((globalThis.cancelAnimationFrame as jest.Mock).mock.calls.length).toBeGreaterThan(0);
    expect(rafCallbacks.size).toBe(0);

    hidden = false;
    act(() => {
      document.dispatchEvent(new Event('visibilitychange'));
    });

    expect(rafCallbacks.size).toBe(1);
    expect(initialQueueSize).toBeGreaterThanOrEqual(1);

    act(() => {
      document.dispatchEvent(new Event('visibilitychange'));
    });

    expect(rafCallbacks.size).toBe(1);

    ReactDom.unmountComponentAtNode(container);
  });
});
