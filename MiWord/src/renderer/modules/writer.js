import { AppModule } from './app-module.js';

/** Bridge between the Office Hub and the existing Writer document engine. */
export class WriterModule extends AppModule {
  async joinCurrentDocument() {
    const collab = this.office?.modules?.collaboration;
    const editor = window.MiWord?.Editor?.el;
    const doc = window.MiWord?.DocumentModel;
    if (!collab || !editor || !doc) return false;
    collab.setToken(localStorage.getItem('tetord.cloud.jwt') || '');
    try { return await collab.joinDocument(String(doc.filePath || doc.id || 'untitled'), 'writer', editor.innerHTML || ''); } catch (_) { return false; }
  }
  newDocument() { this.office?.close(); window.MiWord?.DocumentManager?.newDocument?.(); setTimeout(() => this.joinCurrentDocument(), 0); }
  openNativeRecord(record, d) {
    if (!d || !window.MiWord?.DocumentManager?.confirmDiscard?.()) return false;
    window.MiWord.DocumentModel.replaceDocument({ id: d.id || record.path, name: d.name || record.name, content: d.content || '<p><br></p>', saved: true, sourceType: 'file', fileFormat: d.fileFormat || 'html', headerText: d.headerText || '', footerText: d.footerText || '', showPageNumber: !!d.showPageNumber, differentFirstPage: !!d.differentFirstPage, styles: d.styles });
    window.MiWord.DocumentModel.filePath = record.path;
    window.MiWord.Editor.load(d.content || '<p><br></p>');
    window.MiWord.History.reset(window.MiWord.Editor.el.innerHTML); window.MiWord.PageView.syncDocumentDecorations(); window.MiWord.TitleField.sync();
    this.office?.close(); window.MiWord.Editor.focus(); this.joinCurrentDocument(); return true;
  }
  openDocument(id) {
    const d = window.MiWord?.LocalDocuments?.get?.(id); if (!d || !window.MiWord?.DocumentManager?.confirmDiscard?.()) return false;
    window.MiWord.DocumentModel.replaceDocument({ id: d.id, name: d.name, content: d.content, saved: true, fileFormat: d.fileFormat || 'html', headerText: d.headerText, footerText: d.footerText, showPageNumber: d.showPageNumber, differentFirstPage: d.differentFirstPage, styles: d.styles });
    window.MiWord.Editor.load(d.content); window.MiWord.History.reset(window.MiWord.Editor.el.innerHTML); window.MiWord.PageView.syncDocumentDecorations(); window.MiWord.TitleField.sync();
    this.office?.close(); window.MiWord.Editor.focus(); this.joinCurrentDocument(); return true;
  }
}
