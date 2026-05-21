const fs = require('fs');
const jsdom = require("jsdom");
const { JSDOM } = jsdom;

const html = fs.readFileSync('index.html', 'utf8');
const dom = new JSDOM(html, { runScripts: "outside-only", url: "http://localhost" });
const window = dom.window;
global.window = window;
global.document = window.document;
global.sessionStorage = window.sessionStorage;
global.localStorage = window.localStorage;
global.indexedDB = require("fake-indexeddb");
global.navigator = window.navigator;
global.Event = window.Event;

// Mock fetch if needed
global.fetch = async () => ({ json: async () => ({}) });

const appJs = fs.readFileSync('app.js', 'utf8');
try {
  window.eval(appJs);
  console.log("app.js loaded successfully.");
  
  // Wait for DOMContentLoaded logic to finish
  setTimeout(() => {
    try {
      console.log("Calling renderDashboard...");
      window.renderDashboard();
      console.log("renderDashboard success");
    } catch(e) {
      console.error("renderDashboard Error:", e);
    }
    
    try {
      console.log("Calling renderEmployeeList...");
      window.renderEmployeeList();
      console.log("renderEmployeeList success");
    } catch(e) {
      console.error("renderEmployeeList Error:", e);
    }
  }, 1000);
} catch (e) {
  console.error("Parse error:", e);
}
