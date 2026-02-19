<script lang="ts" context="module">
  export type ExportTarget = "chat" | "team";
</script>

<script lang="ts">
  import { createEventDispatcher } from "svelte";
  import { MessageSquare, Users } from "lucide-svelte";
  import { t } from "../../../i18n/i18n";

  export let target: ExportTarget = "chat";
  export let lang = "en";

  const dispatch = createEventDispatcher<{
    targetChange: ExportTarget;
  }>();

  const targets = [
    { id: "chat" as const, icon: MessageSquare, labelKey: "target.chat" },
    { id: "team" as const, icon: Users, labelKey: "target.team" },
  ];

  const handleChange = (next: ExportTarget) => {
    dispatch("targetChange", next);
  };
</script>

<section class="target-section" data-lang={lang}>
  <div class="card">
    <div class="card-header">
      <div class="card-icon">
        <MessageSquare size={16} />
      </div>
      <h2 class="card-title">{t("target.title", {}, lang)}</h2>
    </div>
    <div
      class="target-toggle"
      role="radiogroup"
      aria-label={t("target.title", {}, lang)}
    >
      {#each targets as item}
        {@const Icon = item.icon}
        <label class="target-pill" class:active={target === item.id}>
          <input
            type="radio"
            name="export-target"
            value={item.id}
            checked={target === item.id}
            on:change={() => handleChange(item.id)}
          />
          <span class="target-pill-icon">
            <Icon size={16} />
          </span>
          <span class="target-pill-label">{t(item.labelKey, {}, lang)}</span>
        </label>
      {/each}
    </div>
  </div>
</section>
