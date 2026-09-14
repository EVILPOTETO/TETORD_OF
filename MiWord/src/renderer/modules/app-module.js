import { storage } from '../core/storage.js';

export class AppModule {
  constructor(ctx = {}) {
    this.office = ctx.office || null;
    this.root = ctx.root || document;
    this.storage = storage;
  }
  $(selector) { return this.root.querySelector(selector); }
  $$(selector) { return [...this.root.querySelectorAll(selector)]; }
  byId(id) { return document.getElementById(id); }
  on(target, event, handler, options) { const el = typeof target === 'string' ? this.byId(target) || this.$(target) : target; if (el) el.addEventListener(event, handler, options); return el; }
  clear(el) { if (el) while (el.firstChild) el.removeChild(el.firstChild); return el; }
  el(tag, { className = '', text = '', attrs = {}, children = [] } = {}) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined && text !== null) node.textContent = String(text);
    Object.entries(attrs).forEach(([key, value]) => { if (value !== undefined && value !== null) node.setAttribute(key, String(value)); });
    children.forEach(child => child && node.appendChild(child));
    return node;
  }
  button(text, className = '', attrs = {}) { return this.el('button', { className, text, attrs: { type: 'button', ...attrs } }); }
  announce(message) { document.dispatchEvent(new CustomEvent('tetord:notify', { detail: { message: String(message) } })); }
}
