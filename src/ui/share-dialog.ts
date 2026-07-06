/**
 * Share dialog controller (W4.4).
 *
 * Manages the native <dialog id="shareDialog"> declared in index.html:
 *  • "Copy link to view" — an OFFLINE deep link (?m=&vp=) via url-state; no
 *    upload needed (works with any already-hosted model URL, or is disabled when
 *    the current model isn't hosted).
 *  • Publish — convert-in-browser is already done (the model is loaded as
 *    fragments); we get its `.frag` buffer and POST it to the hosting API, then
 *    show the embed link, iframe snippet, an offline-generated QR, the expiry,
 *    and a delete-my-upload button (uses the returned delete token).
 *  • PowerPoint tab — Web Viewer add-in steps + the GLB export button.
 *
 * The viewer owns the engine/state; this controller receives callbacks so it
 * stays free of engine internals. All copy is via i18n (t()).
 */
import { formatDate, t } from '../core/i18n';
import { renderQrToCanvas } from '../core/qrcode';
import type { PublishResult } from '../core/upload-client';

/** What the controller needs from the host app. */
export interface ShareDialogCallbacks {
  /** Bytes of the current model's `.frag` + a display file name, or null if none loaded. */
  getFragForShare(): Promise<{ bytes: Uint8Array; fileName: string } | null>;
  /** Publishes bytes → hosting API. Returns the share result. */
  publish(bytes: Uint8Array, fileName: string): Promise<PublishResult>;
  /** Deletes a published upload with its token. */
  deleteUpload(id: string, deleteToken: string): Promise<void>;
  /** Builds the current "copy link to view" deep link (full app, ?m=&vp=), or null. */
  buildCopyLink(): string | null;
  /** Exports the current view as a GLB (the W4.5 path). */
  exportGlb(): Promise<void>;
  /** Toast + status helpers so messages match the rest of the app. */
  toast(message: string, type: 'success' | 'error' | 'info'): void;
  status(message: string): void;
  /** Shows the app's native confirm dialog; resolves true if confirmed. */
  confirm(message: string): Promise<boolean>;
}

interface ShareDom {
  dialog: HTMLDialogElement;
  close: HTMLButtonElement;
  tabLink: HTMLButtonElement;
  tabPp: HTMLButtonElement;
  panelLink: HTMLElement;
  panelPp: HTMLElement;
  copyLink: HTMLButtonElement;
  publish: HTMLButtonElement;
  published: HTMLElement;
  embedUrl: HTMLInputElement;
  copyEmbed: HTMLButtonElement;
  iframe: HTMLTextAreaElement;
  copyIframe: HTMLButtonElement;
  qr: HTMLCanvasElement;
  expiry: HTMLElement;
  delete: HTMLButtonElement;
  hostIntro: HTMLElement;
  ppGlb: HTMLButtonElement;
}

const TTL_DAYS_DISPLAY = 7; // matches the anon entitlement default (display only)

export class ShareDialogController {
  private published: PublishResult | null = null;

  constructor(
    private readonly dom: ShareDom,
    private readonly cb: ShareDialogCallbacks,
  ) {}

  /** Wires all dialog controls once at boot. */
  init(): void {
    this.dom.hostIntro.textContent = t('share.hostIntro', { days: TTL_DAYS_DISPLAY });

    this.dom.close.addEventListener('click', () => this.dom.dialog.close());
    this.dom.tabLink.addEventListener('click', () => this.selectTab('link'));
    this.dom.tabPp.addEventListener('click', () => this.selectTab('pp'));

    this.dom.copyLink.addEventListener('click', () => void this.copyDeepLink());
    this.dom.publish.addEventListener('click', () => void this.publish());
    this.dom.copyEmbed.addEventListener('click', () => void this.copyText(this.dom.embedUrl.value));
    this.dom.copyIframe.addEventListener('click', () => void this.copyText(this.dom.iframe.value));
    this.dom.delete.addEventListener('click', () => void this.deleteUpload());
    this.dom.ppGlb.addEventListener('click', () => void this.cb.exportGlb());

    // Reset language-dependent copy on open.
    this.dom.dialog.addEventListener('close', () => this.reset());
  }

  /** Opens the dialog (native modal — focus trap + Escape handled by <dialog>). */
  open(): void {
    this.dom.hostIntro.textContent = t('share.hostIntro', { days: TTL_DAYS_DISPLAY });
    this.selectTab('link');
    this.dom.dialog.showModal();
  }

  private selectTab(which: 'link' | 'pp'): void {
    const linkActive = which === 'link';
    this.dom.tabLink.classList.toggle('is-active', linkActive);
    this.dom.tabPp.classList.toggle('is-active', !linkActive);
    this.dom.tabLink.setAttribute('aria-selected', String(linkActive));
    this.dom.tabPp.setAttribute('aria-selected', String(!linkActive));
    this.dom.tabLink.tabIndex = linkActive ? 0 : -1;
    this.dom.tabPp.tabIndex = linkActive ? -1 : 0;
    this.dom.panelLink.classList.toggle('is-active', linkActive);
    this.dom.panelPp.classList.toggle('is-active', !linkActive);
    this.dom.panelLink.hidden = !linkActive;
    this.dom.panelPp.hidden = linkActive;
  }

  private async copyDeepLink(): Promise<void> {
    const link = this.cb.buildCopyLink();
    if (!link) {
      this.cb.toast(t('share.needModel'), 'error');
      return;
    }
    await this.copyText(link);
  }

  private async publish(): Promise<void> {
    const frag = await this.cb.getFragForShare();
    if (!frag) {
      this.cb.toast(t('share.needModel'), 'error');
      return;
    }
    this.dom.publish.disabled = true;
    this.dom.publish.textContent = t('share.publishing');
    this.cb.status(t('share.publishing'));
    try {
      const result = await this.cb.publish(frag.bytes, frag.fileName);
      this.published = result;
      this.showResult(result);
      this.cb.status(t('share.published'));
      this.cb.toast(t('share.published'), 'success');
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.cb.toast(t('share.publishFailed', { error: message }), 'error');
      this.cb.status(t('share.publishFailed', { error: message }));
    } finally {
      this.dom.publish.disabled = false;
      this.dom.publish.textContent = t('share.publish');
    }
  }

  private showResult(result: PublishResult): void {
    this.dom.embedUrl.value = result.embedUrl;
    this.dom.iframe.value = buildIframeSnippet(result.embedUrl);
    this.dom.expiry.textContent = t('share.expiresOn', { date: formatDate(result.expiresAt) });
    try {
      renderQrToCanvas(this.dom.qr, result.embedUrl, {
        foreground: '#000000',
        background: '#ffffff',
      });
    } catch {
      // A too-long URL for QR capacity — hide the QR rather than error.
      this.dom.qr.hidden = true;
    }
    this.dom.published.hidden = false;
  }

  private async deleteUpload(): Promise<void> {
    if (!this.published) return;
    const ok = await this.cb.confirm(t('share.deleteConfirm'));
    if (!ok) return;
    try {
      await this.cb.deleteUpload(this.published.id, this.published.deleteToken);
      this.published = null;
      this.dom.published.hidden = true;
      this.dom.qr.hidden = false;
      this.cb.toast(t('share.deleted'), 'success');
      this.cb.status(t('share.deleted'));
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.cb.toast(t('share.deleteFailed', { error: message }), 'error');
    }
  }

  private async copyText(text: string): Promise<void> {
    try {
      await navigator.clipboard.writeText(text);
      this.cb.toast(t('share.linkCopied'), 'success');
    } catch {
      this.cb.toast(t('share.copyFailed'), 'error');
    }
  }

  /** Clears the published result when the dialog closes. */
  private reset(): void {
    this.published = null;
    this.dom.published.hidden = true;
    this.dom.qr.hidden = false;
  }
}

/** Builds the iframe embed snippet for a published embed URL. */
export function buildIframeSnippet(embedUrl: string): string {
  const safe = embedUrl.replace(/"/g, '&quot;');
  return (
    `<iframe src="${safe}" width="800" height="600" frameborder="0" ` +
    `allow="fullscreen" allowfullscreen style="border:0;max-width:100%;" ` +
    `loading="lazy" title="BTC IFC Viewer embed"></iframe>`
  );
}
