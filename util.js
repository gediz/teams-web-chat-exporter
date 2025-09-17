function esc(s) { return (s ?? "").toString().replaceAll('"','""'); }

function toCSV(messages) {
  const header = ['id','author','timestamp','text','reactions','isSystem','replyTo'].join(',');
  const rows = messages.map(m => {
    const reactions = m.reactions?.map(r=>`${r.emoji}:${r.count}`).join('|') ?? '';
    return [m.id, m.author, m.timestamp, m.text?.replaceAll('\n','\\n'), reactions, !!m.isSystem, m.replyTo ?? '']
      .map(v => `"${esc(v)}"`).join(',');
  });
  return [header, ...rows].join('\n');
}

function toMarkdown(messages, meta={}) {
  const head = `# ${meta.title || "Teams Chat Export"}\n\n_Time range_: ${meta.timeRange || "visible"}\n_Participants_: ${meta.participants?.join(", ") || "n/a"}\n\n---\n`;
  const body = messages.map(m => {
    const r = m.reactions?.map(r=>`${r.emoji} ${r.count}`).join(' ') || '';
    const prefix = m.isSystem ? '**[system]**' : `**${m.author}**`;
    const reply = m.replyTo ? `\n> reply to: ${m.replyTo}` : '';
    return `- ${prefix} — _${m.timestamp}_ ${r}\n  \n  ${m.text}${reply}\n`;
  }).join('\n');
  return head + body;
}

function toHTML(messages, meta={}) {
  const style = `<style>body{font:14px system-ui;padding:20px} .msg{margin:8px 0;padding:8px 12px;border:1px solid #eee;border-radius:10px} .sys{background:#f9fafb} .hdr{color:#555;font-size:12px;margin-bottom:6px}</style>`;
  const head = `<h1>${meta.title || "Teams Chat Export"}</h1>
    <p><b>Time range:</b> ${meta.timeRange || "visible"}<br/>
    <b>Participants:</b> ${meta.participants?.join(", ") || "n/a"}</p><hr/>`;
  const body = messages.map(m => {
    const cls = m.isSystem ? "msg sys" : "msg";
    const r = m.reactions?.map(r=>`${r.emoji} ${r.count}`).join(' ');
    const reply = m.replyTo ? `<div class="hdr">reply to: ${m.replyTo}</div>` : '';
    return `<div class="${cls}"><div class="hdr">${m.isSystem ? "[system]" : m.author} — ${m.timestamp} ${r?("— "+r):""}</div><div>${(m.text||"").replaceAll('\n','<br/>')}</div>${reply}</div>`;
  }).join('');
  return `<!doctype html><meta charset="utf-8">${style}${head}${body}`;
}

self.toCSV = toCSV;
self.toMarkdown = toMarkdown;
self.toHTML = toHTML;
