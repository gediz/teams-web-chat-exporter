// Behavior lock for convertApiMessages (src/content/api-converter.ts) — the
// conversion spine that turns the raw Teams API payload into the ExportMessage[]
// every format renders. Fixtures are SYNTHETIC (no real OIDs/thread-ids/names/
// message text) per repo policy. Covers the orchestration convertApiMessages owns:
// per-message conversion, the newest-first -> chronological re-sort, and the
// duplicate-forward-row dedup.
import { test, expect } from 'vitest';
import { convertApiMessages } from '../src/content/api-converter.ts';

const opts = {} as never;
const richText = (o: Record<string, unknown>) => ({ messagetype: 'RichText/Html', ...o } as never);

test('converts a basic rich-text message to author + text', () => {
  const out = convertApiMessages([
    richText({ content: '<div>hello world</div>', imdisplayname: 'Alice', composetime: '2024-01-01T10:00:00Z' }),
  ], opts);
  expect(out).toHaveLength(1);
  expect(out[0].author).toBe('Alice');
  expect(out[0].text).toContain('hello world');
});

test('orders by timestamp even when input array order disagrees (reverse+stable-sort, not just reverse)', () => {
  // Deliberately scrambled: the array order is NOT reverse-chronological, so a plain
  // reverse() would mis-order these (yield first, third, second). Only the stable
  // timestamp sort the converter applies after reverse() produces the right order —
  // this is the re-ranked-tombstone case the sort exists for.
  const out = convertApiMessages([
    richText({ content: '<div>second</div>', imdisplayname: 'A', composetime: '2024-01-01T10:05:00Z' }),
    richText({ content: '<div>third</div>', imdisplayname: 'A', composetime: '2024-01-01T10:10:00Z' }),
    richText({ content: '<div>first</div>', imdisplayname: 'A', composetime: '2024-01-01T10:00:00Z' }),
  ], opts);
  expect(out.map((m) => m.text.trim())).toEqual(['first', 'second', 'third']);
});

test('collapses the duplicate row Teams emits for one forward (same clientmessageid)', () => {
  const fwd = (cid: string) => richText({
    content: '<div>forwarded body</div>',
    imdisplayname: 'Bob',
    composetime: '2024-01-01T10:00:00Z',
    clientmessageid: cid,
    properties: { forwardTemplateId: 'ft1' },
  });
  // Same clientmessageid = Teams' duplicate row for ONE forward -> collapse to one.
  expect(convertApiMessages([fwd('cm1'), fwd('cm1')], opts)).toHaveLength(1);
  // Distinct clientmessageids = two genuine forwards -> both kept.
  expect(convertApiMessages([fwd('cm1'), fwd('cm2')], opts)).toHaveLength(2);
});

test('reads adaptive card text nested in ColumnSet and Container (properties.cards)', () => {
  const card = {
    contentType: 'application/vnd.microsoft.card.adaptive',
    content: { type: 'AdaptiveCard', body: [
      { type: 'ColumnSet', columns: [{ type: 'Column', items: [{ type: 'TextBlock', text: 'Build finished' }] }] },
      { type: 'Container', items: [{ type: 'Container', items: [{ type: 'TextBlock', text: 'All checks passed' }] }] },
    ] },
  };
  const out = convertApiMessages([
    richText({ content: '<div></div>', imdisplayname: 'Build Bot', composetime: '2024-01-01T10:00:00Z', properties: { cards: JSON.stringify([card]) } }),
  ], opts);
  const preview = out[0].attachments?.find((a) => a.kind === 'preview');
  expect(preview?.metaText).toBe('Build finished\nAll checks passed');
});

// RichText/Media_Card as the chat service sends app cards: a "cards.unsupported"
// stub, with the real card base64-encoded in <Swift b64>.
const swiftCard = (cards: unknown[]) => {
  const b64 = Buffer.from(JSON.stringify({ type: 'message/card', attachments: cards }), 'utf8').toString('base64');
  return '<URIObject type="SWIFT.1" url_thumbnail="https://example.invalid/card.png">Card - access it on '
    + '<a href="https://go.skype.com/cards.unsupported">https://go.skype.com/cards.unsupported</a>. '
    + `<Title>Card</Title><Swift b64="${b64}" /><Description /></URIObject>`;
};
// Apps and bots post from a "28:" MRI.
const APP_FROM = 'https://example.invalid/v1/users/ME/contacts/28:00000000-0000-4000-8000-0000000000aa';
const mediaCard = (o: Record<string, unknown>) => richText({ messagetype: 'RichText/Media_Card', imdisplayname: 'Updates', from: APP_FROM, composetime: '2024-01-01T10:00:00Z', ...o });

test('reads a Media_Card (Swift) adaptive card into the message text', () => {
  const card = { contentType: 'application/vnd.microsoft.card.adaptive', content: { type: 'AdaptiveCard', body: [
    { type: 'ColumnSet', columns: [{ type: 'Column', items: [{ type: 'TextBlock', text: 'Daily check-in' }] }] },
    { type: 'Container', items: [
      { type: 'TextBlock', text: 'What did you work on?' },
      { type: 'TextBlock', text: 'Café menu ✅ and résumé page' },
    ] },
    { type: 'ActionSet', actions: [{ type: 'Action.Submit', title: 'Review' }] },
  ] } };
  const out = convertApiMessages([mediaCard({ content: swiftCard([card]) })], opts);
  expect(out[0].text).toBe('Daily check-in\nWhat did you work on?\nCafé menu ✅ and résumé page');
  expect(out[0].text).not.toContain('cards.unsupported');
});

test('a Media_Card whose only text is a collapsible header reads the header', () => {
  const card = { contentType: 'application/vnd.microsoft.card.adaptive', content: { type: 'AdaptiveCard', body: [
    { type: 'Container', selectAction: { type: 'Action.ToggleVisibility', targetElements: ['d'] }, items: [{ type: 'TextBlock', text: 'Deploy succeeded' }] },
  ] } };
  const out = convertApiMessages([mediaCard({ content: swiftCard([card]) })], opts);
  expect(out[0].text).toBe('Deploy succeeded');
});

test('keeps the stub text when the Swift payload does not decode', () => {
  const content = '<URIObject type="SWIFT.1">Card - access it on <a href="https://go.skype.com/cards.unsupported">'
    + 'https://go.skype.com/cards.unsupported</a>. <Title>Card</Title><Swift b64="!!not-base64!!" /></URIObject>';
  const out = convertApiMessages([mediaCard({ content })], opts);
  expect(out[0].text).toContain('cards.unsupported');
});

test('reads Office 365 connector and hero cards sent as Media_Card', () => {
  const connector = { contentType: 'application/vnd.microsoft.teams.card.o365connector', content: {
    summary: 'Build 42 succeeded',
    sections: [{ activityTitle: 'Build 42 succeeded', activitySubtitle: 'main branch', facts: [{ name: 'Duration', value: '3m' }] }],
  } };
  const hero = { contentType: 'application/vnd.microsoft.card.hero', content: { title: 'Release notes', text: '<b>Two</b> fixes' } };
  const summaryOnly = { contentType: 'application/vnd.microsoft.teams.card.o365connector', content: { summary: 'Deploy started' } };
  const out = convertApiMessages([mediaCard({ imdisplayname: 'Build Bot', content: swiftCard([connector, hero, summaryOnly]) })], opts);
  expect(out[0].text).toBe('Build 42 succeeded\nmain branch\nDuration: 3m\n\nRelease notes\nTwo fixes\n\nDeploy started');
});

const PERSON = '00000000-0000-4000-8000-000000000001';
const onBehalfOf = (name: string, uuid: string) => JSON.stringify([{ itemid: 0, mentionType: 'person', mri: `8:orgid:${uuid}`, displayName: name }]);

test('credits the person an app posted for, keeps the app as via, and flags own posts', () => {
  const msg = mediaCard({
    from: 'https://example.invalid/v1/users/ME/contacts/28:00000000-0000-4000-8000-0000000000aa',
    content: '<div>posted by an app</div>',
    properties: { onbehalfof: onBehalfOf('Test Person', PERSON) },
  });
  const [mine] = convertApiMessages([msg], opts, PERSON);
  expect(mine.author).toBe('Test Person');
  expect(mine.via).toBe('Updates');
  expect(mine.isOwn).toBe(true);
  const [theirs] = convertApiMessages([msg], opts, '00000000-0000-4000-8000-000000000002');
  expect(theirs.isOwn).toBeUndefined();
});

test('a normal message has no via and keeps its sender as author', () => {
  const [m] = convertApiMessages([richText({ content: '<div>hi</div>', imdisplayname: 'Alice', composetime: '2024-01-01T10:00:00Z' })], opts);
  expect(m.author).toBe('Alice');
  expect(m.via).toBeUndefined();
});

const cardB64 = (text: string) => Buffer.from(JSON.stringify({ attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', content: { body: [{ type: 'TextBlock', text }] } }] }), 'utf8').toString('base64');
const replyQuote = '<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="m1"><strong itemprop="mri">Updates</strong><p itemprop="preview">Card</p></blockquote>';

test('a Media_Card that quotes a reply keeps the reply and reads the card', () => {
  const content = replyQuote + `<URIObject type="SWIFT.1">Card - access it on <a href="https://go.skype.com/cards.unsupported">x</a>. <Swift b64="${cardB64('Card body')}" /></URIObject>`;
  const [m] = convertApiMessages([mediaCard({ content })], opts);
  expect(m.text).toBe('Card body');
  expect(m.replyTo?.id).toBe('m1');
});

test('a <Swift> element outside an app card does not replace a message body', () => {
  const quoted = `<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="m1"><strong>Updates</strong><URIObject type="SWIFT.1"><Swift b64="${cardB64('quoted card')}" /></URIObject></blockquote><p>my actual reply</p>`;
  const [reply] = convertApiMessages([richText({ content: quoted, imdisplayname: 'Alice', composetime: '2024-01-01T10:00:00Z' })], opts);
  expect(reply.text).toContain('my actual reply');
  expect(reply.replyTo?.author).toBe('Updates');
  const [plain] = convertApiMessages([{ messagetype: 'Text', content: `<Swift b64="${cardB64('injected')}" />`, imdisplayname: 'Alice', composetime: '2024-01-01T10:00:00Z' } as never], opts);
  expect(plain.text).not.toBe('injected');
});

test('one hostile card does not stop the rest of the chat from converting', () => {
  const evil = Buffer.from(JSON.stringify({ attachments: [{ contentType: 'application/vnd.microsoft.teams.card.o365connector', content: { title: { toString: 1 }, sections: [null] } }] }), 'utf8').toString('base64');
  const out = convertApiMessages([
    mediaCard({ content: `<URIObject type="SWIFT.1">stub <Swift b64="${evil}" /></URIObject>`, properties: { onbehalfof: JSON.stringify([{ displayName: { toString: 1 } }]) } }),
    richText({ content: '<div>still here</div>', imdisplayname: 'Alice', composetime: '2024-01-01T10:01:00Z' }),
  ], opts);
  expect(out).toHaveLength(2);
  expect(out[1].text).toContain('still here');
});

test('the person an app posted for names their reactions, and a sender name wins', () => {
  const reaction = (mri: string) => JSON.stringify([{ key: 'like', users: [{ mri, time: 1 }] }]);
  const out = convertApiMessages([
    mediaCard({ content: '<div>check-in</div>', properties: { onbehalfof: onBehalfOf('Test Person', PERSON) } }),
    richText({ content: '<div>nice</div>', imdisplayname: 'Alice', composetime: '2024-01-01T10:01:00Z', properties: { emotions: reaction(`8:orgid:${PERSON}`) } }),
  ], opts);
  expect(out[1].reactions?.[0].reactors?.[0].name).toBe('Test Person');
  const direct = convertApiMessages([
    mediaCard({ content: '<div>check-in</div>', properties: { onbehalfof: onBehalfOf('T. Person', PERSON) } }),
    richText({ content: '<div>hi</div>', imdisplayname: 'Test Person', from: `https://example.invalid/contacts/8:orgid:${PERSON}`, composetime: '2024-01-01T10:01:00Z' }),
    richText({ content: '<div>nice</div>', imdisplayname: 'Alice', composetime: '2024-01-01T10:02:00Z', properties: { emotions: reaction(`8:orgid:${PERSON}`) } }),
  ], opts);
  expect(direct[2].reactions?.[0].reactors?.[0].name).toBe('Test Person');
});

test('a reply quoting an app post names the person it is credited to', () => {
  const out = convertApiMessages([
    mediaCard({ id: 'm1', content: '<div>check-in</div>', properties: { onbehalfof: onBehalfOf('Test Person', PERSON) } }),
    richText({ content: replyQuote + '<p>thanks</p>', imdisplayname: 'Alice', composetime: '2024-01-01T10:01:00Z' }),
  ], opts);
  expect(out[1].replyTo?.author).toBe('Test Person');
});

test('onbehalfof is ignored on forwards and on a person\'s own message, and an app post keeps its via', () => {
  const [fwd] = convertApiMessages([richText({ content: '<div>fwd</div>', imdisplayname: 'Bob', composetime: '2024-01-01T10:00:00Z',
    properties: { forwardTemplateId: 'ft1', onbehalfof: onBehalfOf('Test Person', PERSON) } })], opts);
  expect(fwd.author).not.toBe('Test Person');
  expect(fwd.via).toBeUndefined();
  // A person (8: MRI) cannot credit their own message to someone else.
  const [spoof] = convertApiMessages([richText({ content: '<div>I approve</div>', imdisplayname: 'Mallory', composetime: '2024-01-01T10:00:00Z',
    from: 'https://example.invalid/v1/users/ME/contacts/8:orgid:00000000-0000-4000-8000-0000000000bb', properties: { onbehalfof: onBehalfOf('Test Person', PERSON) } })], opts, PERSON);
  expect(spoof.author).toBe('Mallory');
  expect(spoof.via).toBeUndefined();
  expect(spoof.isOwn).toBeUndefined();
  // An app post always says it came through an app, even when the names match.
  const [same] = convertApiMessages([mediaCard({ imdisplayname: 'Test Person', content: '<div>x</div>', properties: { onbehalfof: onBehalfOf('Test Person', PERSON) } })], opts);
  expect(same.author).toBe('Test Person');
  expect(same.via).toBe('Test Person');
});

test('a card decoded from <Swift> is not repeated as a preview from properties.cards', () => {
  const card = { contentType: 'application/vnd.microsoft.card.adaptive', content: { body: [{ type: 'TextBlock', text: 'Only once' }] } };
  const [m] = convertApiMessages([mediaCard({ content: swiftCard([card]), properties: { cards: JSON.stringify([card]) } })], opts);
  expect(m.text).toBe('Only once');
  expect(m.attachments?.filter((a) => a.kind === 'preview')).toHaveLength(0);
});

test('with replies off, a quoted card never replaces a message\'s own text or card', () => {
  const quotedCard = `<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="m1"><strong>Updates</strong><URIObject type="SWIFT.1"><Swift b64="${cardB64('QUOTED card')}" /></URIObject></blockquote>`;
  const noReplies = { includeReplies: false } as never;
  const [reply] = convertApiMessages([richText({ content: quotedCard + '<p>my actual reply</p>', imdisplayname: 'Alice', composetime: '2024-01-01T10:00:00Z' })], noReplies);
  expect(reply.text).toContain('my actual reply');
  expect(reply.text).not.toContain('QUOTED card');
  const own = `<URIObject type="SWIFT.1"><Swift b64="${cardB64('own card')}" /></URIObject>`;
  const [card] = convertApiMessages([mediaCard({ content: quotedCard + own })], noReplies);
  expect(card.text).toBe('own card');
});

test('only RichText/Media_Card messages are read as app cards', () => {
  const wrapped = `<p>my words</p><URIObject type="SWIFT.1"><Swift b64="${cardB64('injected')}" /></URIObject>`;
  const [html] = convertApiMessages([richText({ content: wrapped, imdisplayname: 'Alice', composetime: '2024-01-01T10:00:00Z' })], opts);
  expect(html.text).toContain('my words');
  const [text] = convertApiMessages([{ messagetype: 'Text', content: `<URIObject type="SWIFT.1"><Swift b64="${cardB64('injected')}" /></URIObject> real text`, imdisplayname: 'Alice', composetime: '2024-01-01T10:00:00Z' } as never], opts);
  expect(text.text).toContain('real text');
});

test('the on-behalf-of name beats a partial mention whatever the message order', () => {
  const reaction = JSON.stringify([{ key: 'like', users: [{ mri: `8:orgid:${PERSON}`, time: 1 }] }]);
  const partialMention = JSON.stringify([{ itemid: 0, mri: `8:orgid:${PERSON}`, displayName: 'Test' }]);
  const msgs = [
    richText({ content: '<div>hi <span itemtype="http://schema.skype.com/Mention" itemid="0">Test</span></div>', imdisplayname: 'Alice', composetime: '2024-01-01T10:02:00Z', properties: { mentions: partialMention, emotions: reaction } }),
    mediaCard({ content: '<div>check-in</div>', composetime: '2024-01-01T10:00:00Z', properties: { onbehalfof: onBehalfOf('Test Person', PERSON) } }),
  ];
  for (const batch of [msgs, [...msgs].reverse()]) {
    const out = convertApiMessages(batch, opts);
    const reacted = out.find((m) => m.reactions?.length);
    expect(reacted?.reactions?.[0].reactors?.[0].name).toBe('Test Person');
  }
});

test('a hostile properties.cards card does not stop the chat from converting', () => {
  const evil = [{ contentType: 'application/vnd.microsoft.card.adaptive', appName: { toString: 1 }, content: {
    selectAction: { url: { toString: 1 } },
    body: [null, { type: 'Image', url: { toString: 1 } }, { type: 'Container', backgroundImage: { url: { toString: 1 } }, items: [null, { type: 'TextBlock', text: 'kept' }] }],
  } }];
  const out = convertApiMessages([
    richText({ content: '<div></div>', imdisplayname: 'Bot', composetime: '2024-01-01T10:00:00Z', properties: { cards: JSON.stringify(evil) } }),
    richText({ content: '<div>still here</div>', imdisplayname: 'Alice', composetime: '2024-01-01T10:01:00Z' }),
  ], opts);
  expect(out).toHaveLength(2);
  expect(out[0].attachments?.find((a) => a.kind === 'preview')?.metaText).toBe('kept');
});

test('a card preview keeps its old title, so it still de-duplicates against a link preview', () => {
  const card = { contentType: 'application/vnd.microsoft.card.adaptive', content: { body: [
    { type: 'ColumnSet', columns: [{ type: 'Column', items: [{ type: 'TextBlock', text: 'Repo name' }] }] },
    { type: 'TextBlock', text: 'Example page' },
  ] } };
  const links = JSON.stringify([{ url: 'https://example.invalid/page', previewenabled: true, preview: { title: 'Example page', previewurl: 'https://example.invalid/page' } }]);
  const [m] = convertApiMessages([richText({ content: '<div>see https://example.invalid/page</div>', imdisplayname: 'Alice', composetime: '2024-01-01T10:00:00Z', properties: { cards: JSON.stringify([card]), links } })], opts);
  const previews = m.attachments?.filter((a) => a.kind === 'preview') ?? [];
  // The card's title is its first top-level TextBlock ("Example page"), as
  // before, so it is folded into the link preview instead of shown twice.
  expect(previews.filter((a) => a.label === 'example.invalid')).toHaveLength(1);
  expect(previews.filter((a) => a.label === 'card')).toHaveLength(0);
});

test('a reply quoting an app post just before the date range still names the person', () => {
  const out = convertApiMessages([
    mediaCard({ id: 'm1', content: '<div>check-in</div>', composetime: '2023-12-31T10:00:00Z', properties: { onbehalfof: onBehalfOf('Test Person', PERSON) } }),
    richText({ content: replyQuote + '<p>thanks</p>', imdisplayname: 'Alice', composetime: '2024-01-01T10:01:00Z' }),
  ], { startAtISO: '2024-01-01T00:00:00Z' } as never);
  expect(out).toHaveLength(1);
  expect(out[0].replyTo?.author).toBe('Test Person');
});

test('a reply quote is matched by server id only, never by an app post\'s clientmessageid', () => {
  const out = convertApiMessages([
    mediaCard({ id: undefined, clientmessageid: 'm1', content: '<div>check-in</div>', properties: { onbehalfof: onBehalfOf('Test Person', PERSON) } }),
    richText({ content: replyQuote + '<p>thanks</p>', imdisplayname: 'Alice', composetime: '2024-01-01T10:01:00Z' }),
  ], opts);
  expect(out.find((m) => m.author === 'Alice')?.replyTo?.author).toBe('Updates');
});

test('a Media_Card ignores cards inside any quote, not just the first reply quote', () => {
  const stub = (b64: string) => `<URIObject type="SWIFT.1">Card - access it on <a href="https://go.skype.com/cards.unsupported">x</a>. <Swift b64="${b64}" /></URIObject>`;
  const second = replyQuote.replace('itemid="m1"', 'itemid="m2"').replace('<p itemprop="preview">Card</p>', stub(cardB64('Second quote')));
  const plain = `<BLOCKQUOTE>${stub(cardB64('Plain quote'))}</BLOCKQUOTE>`;
  for (const content of [replyQuote + second + stub(cardB64('Own card')), plain + stub(cardB64('Own card'))]) {
    for (const includeReplies of [true, false]) {
      const [m] = convertApiMessages([mediaCard({ content })], { ...opts, includeReplies } as never);
      expect(m.text).toBe('Own card');
    }
  }
});

test('an audio card whose url or duration is not a string does not break the export', () => {
  const [m] = convertApiMessages([richText({ content: '<p>voice</p>', properties: { cards: JSON.stringify([
    { contentType: 'application/vnd.microsoft.card.audio', content: { duration: { x: 1 }, media: [{ url: { a: 1 } }] } },
    { contentType: 'application/vnd.microsoft.card.audio', content: { duration: 'PT5S', media: [{ url: 'https://example.invalid/v.mp4' }] } },
  ]) } })], opts);
  const audio = m.attachments?.filter((a) => a.type === 'audio') ?? [];
  expect(audio.map((a) => a.href)).toEqual(['https://example.invalid/v.mp4']);
});

test('the app name in via is flattened to one line too', () => {
  const [m] = convertApiMessages([mediaCard({ imdisplayname: 'Updates\n[10:01] Boss: fake', content: '<div>x</div>', properties: { onbehalfof: onBehalfOf('Test Person', PERSON) } })], opts);
  expect(m.via).toBe('Updates [10:01] Boss: fake');
});
