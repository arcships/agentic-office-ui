<template>
  <div class="app">
    <nav class="nav">
      <h1 class="nav-title">Vue Extend UI</h1>
      <div class="nav-links">
        <router-link to="/">🏠 Home</router-link>
        <router-link to="/docx-viewer">📄 DOCX Viewer</router-link>
        <router-link to="/docx-editor">✏️ DOCX Editor</router-link>
        <router-link to="/xlsx-viewer">📊 XLSX Viewer</router-link>
        <router-link to="/pdf-viewer">📑 PDF Viewer</router-link>
        <router-link to="/components">🧩 Components</router-link>
      </div>
    </nav>
    <main class="main">
      <router-view />
    </main>
  </div>
</template>

<style scoped>
.app { min-height: 100vh; display: flex; flex-direction: column; width: 100%; min-width: 0; overflow-x: hidden; }
.nav {
  display: flex; align-items: center; gap: 24px; flex-wrap: wrap;
  padding: 12px 24px; border-bottom: 1px solid var(--border);
  background: var(--muted); min-width: 0;
}
.nav-title { font-size: 18px; font-weight: 700; flex: 0 0 auto; }
.nav-links { display: flex; gap: 16px; flex-wrap: wrap; min-width: 0; }
.nav-links a { color: var(--foreground); text-decoration: none; font-size: 14px; white-space: nowrap; }
.nav-links a:hover, .nav-links a.router-link-active { color: var(--primary); font-weight: 600; }
.main { flex: 1; display: flex; flex-direction: column; min-width: 0; overflow-x: hidden; }
@media (max-width: 640px) {
  .nav { gap: 10px; padding: 10px 12px; align-items: flex-start; }
  .nav-title { width: 100%; }
  .nav-links { gap: 8px 12px; }
  .nav-links a { font-size: 13px; }
}
</style>
