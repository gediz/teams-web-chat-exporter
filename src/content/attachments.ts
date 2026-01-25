import type { Attachment } from '../types/shared';
import { textFrom } from '../utils/text';

export async function extractAttachments(item: Element, body: Element): Promise<Attachment[]> {
  const map = new Map<string, Attachment>();
  const pending: Promise<void>[] = [];
  const fetchCache = new Map<string, Promise<string | null>>();
  const merge = (prev: Attachment, next: Attachment): Attachment => {
    const merged: Attachment = { ...prev };
    if (!merged.href && next.href) merged.href = next.href;
    if (!merged.label && next.label) merged.label = next.label;
    for (const field of ['type', 'size', 'owner', 'metaText', 'dataUrl'] as const) {
      const val = next[field];
      if (!merged[field] && val) merged[field] = val;
    }
    return merged;
  };
  const guessTypeFromLabel = (label = '') => {
    const m = label.trim().match(/\.([A-Za-z0-9]{1,6})$/);
    return m ? m[1].toUpperCase() : null;
  };
  const collectMetaText = (node: Element | null, label?: string) => {
    const parts = new Set<string>();
    const add = (val?: string | null) => {
      if (!val || typeof val !== 'string') return;
      const trimmed = val.trim();
      if (!trimmed) return;
      parts.add(trimmed);
    };
    if (node) {
      add(node.getAttribute?.('aria-label'));
      add(node.getAttribute?.('title'));
      const txt = node.textContent?.trim();
      if (txt && txt !== label?.trim()) add(txt);
    }
    if (!parts.size) return '';
    return Array.from(parts).join(' • ').replace(/\s+/g, ' ').trim();
  };
  const inferOwner = (text = '') => {
    const match = text.match(/(?:Shared by|Uploaded by|Sent by|From|Owner)\s*:?\s*([^•]+)/i);
    return match ? match[1].trim() : null;
  };
  const inferSize = (text = '') => {
    const match = text.match(/\b\d+(?:[.,]\d+)?\s*(?:bytes?|KB|MB|GB|TB)\b/i);
    return match ? match[0].replace(',', '.').trim() : null;
  };
  const inferType = (label: string, text?: string | null) => {
    return (
      guessTypeFromLabel(label) ||
      (text ? text.match(/\b(PDF|DOCX|XLSX|PPTX|TXT|PNG|JPE?G|GIF|ZIP|RAR|CSV|MP4|MP3)\b/i)?.[0]?.toUpperCase() || null : null)
    );
  };
  const push = (sourceNode: Element | null, data: Partial<Attachment> = {}) => {
    const att: Attachment = { ...data };
    const linkish = sourceNode as HTMLAnchorElement | null;
    if (!att.href && linkish?.href) att.href = linkish.href;
    if (!att.label) {
      const ariaLabel = sourceNode?.getAttribute?.('aria-label');
      if (ariaLabel) att.label = ariaLabel.split(/\n+/)[0].trim();
    }
    if (!att.label && sourceNode?.getAttribute?.('title')) {
      att.label = sourceNode.getAttribute('title')!.split(/\n+/)[0].trim();
    }
    if (!att.label && sourceNode?.textContent) {
      const text = sourceNode.textContent.trim();
      if (text) att.label = text.split(/\n+/)[0].trim();
    }
    if (!att.href && !att.label) return null;

    const metaText = collectMetaText(sourceNode, att.label || undefined);
    if (metaText) att.metaText = metaText;
    const type = inferType(att.label || '', metaText);
    if (type) att.type = type;
    const size = inferSize(metaText || '');
    if (size) att.size = size;
    const owner = inferOwner(metaText || '');
    if (owner) att.owner = owner;

    const key = `${att.href || ''}@@${att.label || ''}`;
    const prev = map.get(key);
    const stored = prev ? merge(prev, att) : att;
    map.set(key, stored);
    return stored;
  };
  const imageToDataUrl = (img: HTMLImageElement | null) => {
    if (!img) return null;
    const src = img.getAttribute('src') || '';
    if (!src.startsWith('blob:') && !src.startsWith('data:')) return null;
    if (!img.complete || !img.naturalWidth || !img.naturalHeight) return null;
    try {
      const canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      const ctx = canvas.getContext('2d');
      if (!ctx) return null;
      ctx.drawImage(img, 0, 0);
      return canvas.toDataURL('image/png');
    } catch {
      return null;
    }
  };
  const isSmallDataUrl = (dataUrl: string | null) => {
    if (!dataUrl) return true;
    const idx = dataUrl.indexOf(',');
    const b64 = idx >= 0 ? dataUrl.slice(idx + 1) : dataUrl;
    return b64.length < 4096;
  };
  const fetchImageAsDataUrl = async (url: string) => {
    if (!/^https?:\/\//i.test(url)) return null;
    let cached = fetchCache.get(url);
    if (!cached) {
      cached = (async () => {
        try {
          const res = await fetch(url, { credentials: 'include' });
          if (!res.ok) return null;
          const buf = await res.arrayBuffer();
          const bytes = new Uint8Array(buf);
          let bin = '';
          for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
          const b64 = btoa(bin);
          const ct = res.headers.get('content-type') || 'image/png';
          return `data:${ct};base64,${b64}`;
        } catch {
          return null;
        }
      })();
      fetchCache.set(url, cached);
    }
    return cached;
  };
  const pickLargestImage = (images: HTMLImageElement[]) => {
    let best: HTMLImageElement | null = null;
    let bestArea = 0;
    for (const img of images) {
      const area = (img.naturalWidth || 0) * (img.naturalHeight || 0);
      if (area > bestArea) {
        bestArea = area;
        best = img;
      }
    }
    return best;
  };
  const extractAdaptiveCardPreview = (root: Element) => {
    const findLinkIn = (node: Element | null) => {
      if (!node) return '';
      const linkEl =
        node.querySelector<HTMLAnchorElement>('a[data-testid="atp-safelink"][href^="http"]') ||
        node.querySelector<HTMLAnchorElement>('a[href^="http"]');
      if (linkEl?.href) return linkEl.href;
      const ariaLabel = node.getAttribute('aria-label') || '';
      const match = ariaLabel.match(/https?:\/\/\S+/i);
      return match ? match[0] : '';
    };

    const cards = root.querySelectorAll<HTMLElement>('[data-tid="adaptive-card"], .ac-adaptiveCard');
    if (!cards.length) return;

    cards.forEach(card => {
      const container = card.closest('[aria-label*="card message"]') || root;
      const linkHref = findLinkIn(container) || findLinkIn(root);
      let host = '';
      if (linkHref) {
        try {
          host = new URL(linkHref).hostname;
        } catch {
          host = '';
        }
      }

      const blocks = Array.from(card.querySelectorAll<HTMLElement>('.ac-textBlock'));
      const lines = blocks
        .map(b => textFrom(b))
        .map(s => s.replace(/\s+/g, ' ').trim())
        .filter(Boolean);

      const imgCandidates = Array.from(card.querySelectorAll<HTMLImageElement>('img'));
      const bestImg = pickLargestImage(imgCandidates);
      const dataUrl = imageToDataUrl(bestImg);
      let imgSrc = dataUrl || bestImg?.getAttribute('src') || '';
      if (imgSrc.startsWith('blob:') && !dataUrl) imgSrc = '';

      if (!lines.length && !imgSrc) return;

      const label = host ? `${host} preview` : 'link preview';
      const metaText = lines.join('\n');
      push(card, { href: imgSrc || undefined, label, metaText, kind: 'preview' });
    });
  };
  const parseTitle = (t: string | null) => {
    if (!t) return null;
    const parts = t.split(/\n+/).map(s => s.trim()).filter(Boolean);
    if (parts.length >= 2 && /^https?:\/\//i.test(parts[1])) {
      return { label: parts[0], href: parts[1], metaText: parts.slice(2).join(' • ') };
    }
    const m = t.match(/https?:\/\/\S+/);
    if (m) {
      const url = m[0];
      const label = parts[0] && parts[0] !== url ? parts[0] : url;
      return { label, href: url, metaText: parts.slice(1).join(' • ') };
    }
    return null;
  };

  const roots: Element[] = [];
  const aria = body?.getAttribute('aria-labelledby') || '';
  const attId = aria.split(/\s+/).find(s => s.startsWith('attachments-'));
  if (attId) {
    const el = document.getElementById(attId);
    if (el) roots.push(el);
  }
  ['[data-tid="file-attachment-grid"]', '[data-tid="file-preview-root"]', '[data-tid="attachments"]'].forEach(sel => {
    const el = body && body.querySelector(sel);
    if (el && !roots.includes(el)) roots.push(el);
  });

  for (const root of roots) {
    root.querySelectorAll<HTMLElement>('[data-tid="file-preview-root"][amspreviewurl]').forEach(node => {
      const href = node.getAttribute('amspreviewurl') || '';
      if (!href) return;
      const previewImg = node.querySelector<HTMLImageElement>('img[data-tid="rich-file-preview-image"]');
      const dataUrl = imageToDataUrl(previewImg);
      const label =
        node.getAttribute('aria-label') ||
        node.getAttribute('title') ||
        'image';
      push(node, { href, label, dataUrl: dataUrl || undefined });
    });
    root.querySelectorAll('[data-testid="file-attachment"], [data-tid^="file-chiclet-"]').forEach(el => {
      const t = el.getAttribute('title') || el.getAttribute('aria-label') || '';
      const parsed = parseTitle(t);
      if (parsed) push(el, parsed);
      el.querySelectorAll<HTMLAnchorElement>('a[href^="http"]').forEach(a => {
        const label = textFrom(a) || a.getAttribute('aria-label') || a.title || a.href;
        push(a, { href: a.href, label });
      });
    });
    root.querySelectorAll('button[data-testid="rich-file-preview-button"][title]').forEach(btn => {
      const parsed = parseTitle(btn.getAttribute('title'));
      if (parsed) push(btn, parsed);
    });
    root.querySelectorAll<HTMLAnchorElement>('a[href^="http"]').forEach(a => {
      const label = textFrom(a) || a.getAttribute('aria-label') || a.title || a.href;
      push(a, { href: a.href, label });
    });
  }

  const contentRoot = body && (body.querySelector('[id^="content-"]') || body.querySelector('[data-tid="message-content"]'));
  if (contentRoot) {
    contentRoot.querySelectorAll<HTMLAnchorElement>('a[data-testid="atp-safelink"], a[href^="http"]').forEach(a => {
      push(a, { href: a.href, label: textFrom(a) || a.getAttribute('aria-label') || a.title || a.href });
    });

    // OLD:
    // contentRoot.querySelectorAll<HTMLImageElement>('[data-testid="lazy-image-wrapper"] img').forEach(img => {
    //   const src = img.getAttribute('src') || '';
    //   if (/^https?:\/\//i.test(src)) {
    //     push(img, { href: src, label: img.getAttribute('alt') || 'image' });
    //   }
    // });

    // NEW: inline AMS images + file preview AMS URLs (full-size)
    contentRoot
      .querySelectorAll<HTMLImageElement>('span[itemtype="http://schema.skype.com/AMSImage"] img[data-gallery-src], img[itemtype="http://schema.skype.com/AMSImage"][data-gallery-src]')
      .forEach(img => {
        const gallery = img.getAttribute('data-gallery-src') || '';
        const orig = img.getAttribute('data-orig-src') || '';
        const href = gallery || orig;
        if (!href) return;

        const label =
          img.getAttribute('alt') ||
          img.getAttribute('aria-label') ||
          img.getAttribute('title') ||
          'image';

        const dataUrl = imageToDataUrl(img);
        // This will build an Attachment { href, label, type/size/owner/metaText inferred }
        const stored = push(img, { href, label, dataUrl: dataUrl || undefined });
        if (stored && isSmallDataUrl(dataUrl)) {
          pending.push(
            fetchImageAsDataUrl(href).then(fetched => {
              if (!fetched) return;
              if (!stored.dataUrl || isSmallDataUrl(stored.dataUrl)) {
                stored.dataUrl = fetched;
              }
            }),
          );
        }
      });

    extractAdaptiveCardPreview(contentRoot);
  }

  if (pending.length) {
    await Promise.all(pending);
  }
  return Array.from(map.values());
}
