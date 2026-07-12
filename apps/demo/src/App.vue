<template>
  <div class="app">
    <nav class="nav">
      <h1 class="nav-title">Vue Extend UI</h1>
      <div class="nav-links">
        <router-link to="/">Home</router-link>
        <router-link to="/docx-surface">DOCX Surface</router-link>
        <router-link to="/docx-viewer">DOCX Viewer</router-link>
        <router-link to="/docx-editor">DOCX Editor</router-link>
        <router-link to="/xlsx-surface">XLSX Surface</router-link>
        <router-link to="/xlsx-viewer">XLSX Viewer</router-link>
        <router-link to="/pptx-viewer">PPTX Viewer</router-link>
        <router-link to="/pptx-headless">PPTX 最小组合</router-link>
        <router-link to="/pdf-viewer">PDF Viewer</router-link>
        <router-link to="/components">Components</router-link>
      </div>
    </nav>
    <main class="main">
      <router-view />
    </main>
  </div>
</template>

<style scoped>
.app { min-height: 100vh; display: flex; flex-direction: column; width: 100%; min-width: 0; overflow-x: hidden; background: var(--background); }
.nav {
  display: flex; align-items: center; gap: 20px; flex-wrap: wrap;
  padding: 10px 24px; border-bottom: 1px solid var(--border);
  background: color-mix(in oklch, var(--background) 88%, var(--muted)); min-width: 0;
  position: sticky; top: 0; z-index: 20; backdrop-filter: blur(8px);
}
.nav-title { font-size: 15px; font-weight: 650; flex: 0 0 auto; letter-spacing: -0.01em; }
.nav-links { display: flex; gap: 4px; flex-wrap: wrap; min-width: 0; }
.nav-links a {
  color: var(--muted-foreground); text-decoration: none; font-size: 13px; white-space: nowrap;
  border-radius: var(--radius-md); padding: 6px 10px; line-height: 1.25;
}
.nav-links a:hover { color: var(--foreground); background: var(--accent); }
.nav-links a.router-link-active { color: var(--foreground); background: var(--secondary); font-weight: 600; }
.main { flex: 1; display: flex; flex-direction: column; min-width: 0; overflow-x: hidden; }
@media (max-width: 640px) {
  .nav { gap: 8px; padding: 10px 12px; align-items: flex-start; }
  .nav-title { width: 100%; }
  .nav-links { gap: 4px; }
  .nav-links a { font-size: 12px; padding: 5px 8px; }
}
</style>
