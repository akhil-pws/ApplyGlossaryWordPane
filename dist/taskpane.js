!function(){"use strict";var e,t,n,o={14385:function(e){e.exports=function(e,t){return t||(t={}),e?(e=String(e.__esModule?e.default:e),t.hash&&(e+=t.hash),t.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(e)?'"'.concat(e,'"'):e):e}},58394:function(e,t,n){e.exports=n.p+"be5043e12582d7bc2117.css"}},a={};function r(e){var t=a[e];if(void 0!==t)return t.exports;var n=a[e]={exports:{}};return o[e](n,n.exports,r),n.exports}r.m=o,r.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return r.d(t,{a:t}),t},r.d=function(e,t){for(var n in t)r.o(t,n)&&!r.o(e,n)&&Object.defineProperty(e,n,{enumerable:!0,get:t[n]})},r.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),r.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;r.g.importScripts&&(e=r.g.location+"");var t=r.g.document;if(!e&&t&&(t.currentScript&&(e=t.currentScript.src),!e)){var n=t.getElementsByTagName("script");if(n.length)for(var o=n.length-1;o>-1&&(!e||!/^http(s?):/.test(e));)e=n[o--].src}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),r.p=e}(),r.b=document.baseURI||self.location.href,function(){let e="",t="",n=[],o=!0,a=[],r="",s=!1,i="",c=[],l=[],d=!1;function p(){document.getElementById("app-body").innerHTML='\n    <div class="container mt-5">\n      <form id="login-form" class="p-4 border rounded">\n        <div class="mb-3">\n          <label for="organization" class="form-label fw-bold">Organization</label>\n          <input type="text" class="form-control" id="organization" required>\n        </div>\n        <div class="mb-3">\n          <label for="username" class="form-label fw-bold">Username</label>\n          <input type="text" class="form-control" id="username" required>\n        </div>\n        <div class="mb-3">\n          <label for="password" class="form-label fw-bold">Password</label>\n          <input type="password" class="form-control" id="password" required>\n        </div>\n        <div class="d-grid">\n          <button type="submit" class="btn btn-primary">Login</button>\n        </div>\n      </form>\n    </div>\n  ',document.getElementById("login-form").addEventListener("submit",m)}async function m(t){t.preventDefault();const n=document.getElementById("organization").value,o=document.getElementById("username").value,a=document.getElementById("password").value;document.getElementById("app-body").innerHTML='\n  <div id="button-container">\n\n          <div class="loader" id="loader"></div>\n          </div\n';try{const t=await fetch("https://plsdevapp.azurewebsites.net/api/user/login",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({ClientName:n,Username:o,Password:a})});if(!t.ok)throw p(),new Error("Network response was not ok.");const r=await t.json();!0===r.Status&&r.Data&&r.Data.ResponseStatus?(e=r.Data.Token,sessionStorage.setItem("token",e),window.location.hash="#/dashboard"):p()}catch(e){p(),console.error("Error during login:",e)}}async function u(t,n){if(t.FilteredReportHeadAIHistoryList&&0!==t.FilteredReportHeadAIHistoryList.length||await async function(t){try{const n=await fetch(`https://plsdevapp.azurewebsites.net/api/report/ai-history/${t.ID}`,{method:"GET",headers:{Authorization:`Bearer ${e}`}});if(!n.ok)throw new Error("Network response was not ok.");const o=await n.json();return t.ReportHeadAIHistoryList=o.Data||[],t.FilteredReportHeadAIHistoryList=[],t.ReportHeadAIHistoryList.forEach(((e,n)=>{e.Response=x(e.Response),t.FilteredReportHeadAIHistoryList.unshift(e)})),t.FilteredReportHeadAIHistoryList}catch(e){return console.error("Error fetching AI history:",e),[]}}(t),t.FilteredReportHeadAIHistoryList.length>0){const o=t.FilteredReportHeadAIHistoryList.map(((e,t)=>`<div class="row chatbox">\n        <div class="col-md-12 mt-2 p-2">\n          <span class="ms-3">\n            <i class="fa fa-copy text-secondary c-pointer" title="Copy Response" id="copyPrompt-${n}-${t}"></i>\n          </span>\n          <span class="float-end w-75 me-3">\n            <div class="form-control h-34 d-flex align-items-center dynamic-height user">\n              ${e.Prompt}\n            </div>\n          </span>\n        </div>\n        <div class="col-md-12 mt-2 p-2 d-flex align-items-center">\n          <span class="radio-select">\n            <input class="form-check-input c-pointer" type="radio" name="flexRadioDefault-${n}"\n              id="flexRadioDefault1-${n}-${t}" ${1===e.Selected?"checked":""}>\n          </span>\n          <span class="ms-2 w-75">\n            <div class="form-control h-34 d-flex align-items-center dynamic-height ai-reply ${1===e.Selected?"ai-selected-reply":"bg-light"}" id='selected-response-${n}${t}'>\n              ${e.Response}\n            </div>\n          </span>\n          <span class="ms-2">\n            <i class="fa fa-copy text-secondary c-pointer" title="Copy Response" id="copyResponse-${n}-${t}"></i>\n          </span>\n        </div>\n\n\n      </div>`)).join("");return setTimeout((()=>{t.FilteredReportHeadAIHistoryList.forEach(((o,a)=>{document.getElementById(`copyPrompt-${n}-${a}`)?.addEventListener("click",(()=>f(o.Prompt))),document.getElementById(`copyResponse-${n}-${a}`)?.addEventListener("click",(()=>f(o.Response))),document.getElementById(`flexRadioDefault1-${n}-${a}`)?.addEventListener("change",(()=>async function(t,n,o){if(!d){d=!0,document.getElementById(`sendPrompt-${n}`).innerHTML='<i class="fa fa-spinner fa-spin text-white"></i>';const a=t.FilteredReportHeadAIHistoryList[o];let r=JSON.parse(JSON.stringify(a));r.Container=l.Container,r.Selected=1;try{const s=await fetch("https://plsdevapp.azurewebsites.net/api/report/ai-history/update",{method:"PUT",headers:{"Content-Type":"application/json",Authorization:`Bearer ${e}`},body:JSON.stringify(r)});if(!s.ok)throw new Error("Network response was not ok.");const i=await s.json();if(i.Data){t.ReportHeadAIHistoryList=JSON.parse(JSON.stringify(i.Data)),t.FilteredReportHeadAIHistoryList=[],t.ReportHeadAIHistoryList.forEach((e=>{e.Response=x(e.Response),t.FilteredReportHeadAIHistoryList.unshift(e)}));document.getElementById(`selected-response-parent-${n}`).querySelectorAll(".ai-selected-reply").forEach((e=>{e.classList.remove("ai-selected-reply"),e.classList.add("bg-light")}));const e=document.getElementById(`selected-response-${n}${o}`);e&&(e.classList.remove("bg-light"),e.classList.add("ai-selected-reply")),t.UserValue=a.Response,t.EditorValue=a.Response,t.text=a.Response}}catch(e){console.error("Error updating AI data:",e)}finally{document.getElementById(`sendPrompt-${n}`).innerHTML='<i class="fa fa-paper-plane text-white"></i>',d=!1}}}(t,n,a)))}))}),0),o}return"<div>No AI history available.</div>"}function y(e,t,n,o,a){return`\n   <h2 class="accordion-header" id="${e}">\n  <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse"\n    data-bs-target="#${t}" aria-expanded="false" aria-controls="${t}">\n    ${n.DisplayName}\n  </button>\n</h2>\n<div id="${t}" class="accordion-collapse collapse" aria-labelledby="${e}">\n  <div class="accordion-body chatbox" id="selected-response-parent-${a}">\n    ${o}\n  </div>\n\n  <div class="col-md-12 d-flex align-items-center justify-content-end chatbox p-3">\n    <textarea class="form-control" rows="3" id="chatbox-${a}" placeholder="Type here"></textarea>\n    <div class="d-flex align-self-end">\n      <button type="submit" class="btn btn-primary ms-2 text-white" id="sendPrompt-${a}">\n        <i class="fa fa-paper-plane text-white"></i>\n      </button>\n    </div>\n  </div>\n</div>\n  `}async function h(t,n,o){if(""===n||d)console.error("No empty prompt allowed");else{d=!0,document.getElementById(`sendPrompt-${o}`).innerHTML='<i class="fa fa-spinner fa-spin text-white"></i>';const a={ReportHeadID:t.FilteredReportHeadAIHistoryList[0].ReportHeadID,DocumentID:l.NCTID,DocumentType:l.DocumentType,TextSetting:l.TextSetting,DocumentTemplate:l.ReportTemplate,ReportHeadGroupKeyID:t.FilteredReportHeadAIHistoryList[0].ReportHeadGroupKeyID,Container:l.Container,GroupName:i,Prompt:n,PromptType:1,Response:"",Selected:0,ID:0};try{const n=await fetch("https://plsdevapp.azurewebsites.net/api/report/ai-history/add",{method:"POST",headers:{"Content-Type":"application/json",Authorization:`Bearer ${e}`},body:JSON.stringify(a)});if(!n.ok)throw new Error("Network response was not ok.");const r=await n.json();if(r.Data&&"false"!==r.Data){t.ReportHeadAIHistoryList=JSON.parse(JSON.stringify(r.Data)),t.FilteredReportHeadAIHistoryList=[],t.ReportHeadAIHistoryList.forEach(((e,n)=>{e.Response=x(e.Response),t.FilteredReportHeadAIHistoryList.unshift(e)}));const e=`flush-collapseOne-${o}`,n=`flush-headingOne-${o}`,a=document.getElementById(`accordion-item-${o}`),s=await u(t,o);a.innerHTML=y(n,e,t,s,o),document.getElementById(`sendPrompt-${o}`).innerHTML='<i class="fa fa-paper-plane text-white"></i>',d=!1}else document.getElementById(`sendPrompt-${o}`).innerHTML='<i class="fa fa-paper-plane text-white"></i>',d=!1}catch(e){document.getElementById(`sendPrompt-${o}`).innerHTML='<i class="fa fa-paper-plane text-white"></i>',d=!1,console.error("Error sending ai prompt:",e)}}}function f(e){const t=document.createElement("textarea");t.value=e,document.body.appendChild(t),t.select(),document.execCommand("copy"),document.body.removeChild(t)}async function g(){s&&await L(),document.getElementById("app-body").innerHTML='\n  <div class="d-flex justify-content-end p-1">\n     <button class="btn btn-primary btn-sm c-pointer text-white me-2 mb-2" id="applyAITag">\n        <i class="fa fa-robot text-light"></i>\n        Apply\n    </button>\n    </div>\n\n    <div class="card-container"  id="card-container">\n    </div>\n  ';const e=document.getElementById("card-container");document.getElementById("applyAITag").addEventListener("click",b);for(let t=0;t<n.length;t++){const o=n[t],a=document.createElement("div");a.classList.add("accordion-item"),a.id=`accordion-item-${t}`;const r=`flush-headingOne-${t}`,s=`flush-collapseOne-${t}`,i=await u(o,t);a.innerHTML=y(r,s,o,i,t),e.appendChild(a),document.getElementById(`sendPrompt-${t}`)?.addEventListener("click",(()=>{const e=document.getElementById(`chatbox-${t}`).value;h(o,e,t)}))}document.querySelectorAll(".accordion-button").forEach((e=>{e.addEventListener("click",(function(){const e=this.nextElementSibling;e&&e.classList&&e.classList.toggle("show")}))})),document.querySelectorAll(".fa-copy").forEach((e=>{e.addEventListener("click",(function(){this.closest(".p-2").querySelector(".form-control").textContent}))}))}async function b(){return Word.run((async e=>{try{const t=e.document.body;e.load(t,"text"),await e.sync();for(let o=0;o<n.length;o++){const a=n[o];a.EditorValue=x(a.EditorValue);const r=t.search(`#${a.DisplayName}#`,{matchCase:!0,matchWholeWord:!0});e.load(r,"items"),await e.sync(),console.log(`Found ${r.items.length} instances of #${a.DisplayName}#`),r.items.forEach((e=>{""!==a.EditorValue&&e.insertText(a.EditorValue,Word.InsertLocation.replace)})),await e.sync()}await e.sync()}catch(e){console.error("Error during tag application:",e)}}))}async function w(){if(!d)if(0===c.length){document.getElementById("app-body").innerHTML='\n  <div id="button-container">\n\n          <div class="loader" id="loader"></div>\n\n        <div id="highlighted-text"></div>';try{const t=await fetch("https://plsdevapp.azurewebsites.net/api/glossary-template/id/3",{method:"GET",headers:{Authorization:`Bearer ${e}`}});if(!t.ok)throw new Error("Network response was not ok.");const n=await t.json();c=n.Data.GlossaryTemplateData,r=n.Data.Name,E()}catch(e){console.error("Error fetching glossary data:",e)}}else E()}function E(){document.getElementById("app-body").innerHTML='\n        <div id="button-container">\n          <button class="btn btn-secondary me-2 mark-glossary btn-sm" id="applyglossary">Apply Glossary</button>\n        </div>\n  ',document.getElementById("applyglossary").addEventListener("click",v)}async function v(){document.getElementById("app-body").innerHTML='\n  <div id="button-container">\n\n          <div class="loader" id="loader"></div>\n\n        <div id="highlighted-text"></div>';try{await Word.run((async e=>{const t=e.document.body,n=c.map((e=>{const n=t.search(e.ClinicalTerm,{matchCase:!0,matchWholeWord:!0});return n.load("items"),n}));await e.sync(),n.forEach((e=>{e.items.forEach((e=>{e.font.highlightColor="yellow"}))})),s=!0,document.getElementById("app-body").innerHTML='\n      <div id="button-container">\n        <button class="btn btn-secondary me-2 clear-glossary btn-sm" id="clearGlossary">Clear Glossary</button>\n      </div>\n\n      <div id="highlighted-text"></div>\n      \n',await e.sync(),document.getElementById("clearGlossary").addEventListener("click",L),Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged,I)})),console.log("Glossary applied successfully")}catch(e){console.error("Error applying glossary:",e),console.log("Error applying glossary. Please try again.")}}async function I(){await async function(){try{await Word.run((async e=>{const t=e.document.getSelection();if(t.load("text, font.highlightColor"),await e.sync(),t.text){const n=c.map((e=>{const n=t.search(e.ClinicalTerm,{matchCase:!1,matchWholeWord:!0});return n.load("items"),n}));await e.sync();const o=[];n.forEach((e=>{e.items.forEach((e=>{o.push(e.text)}))})),function(e){const t=document.getElementById("highlighted-text");if(t){t.innerHTML="";const n={};e.forEach((e=>{c.forEach((t=>{t.ClinicalTerm===e&&(n[t.ClinicalTerm]||(n[t.ClinicalTerm]=[]),n[t.ClinicalTerm].includes(t.LayTerm)||n[t.ClinicalTerm].push(t.LayTerm))}))})),Object.keys(n).forEach((e=>{const o=document.createElement("div");o.className="box";const a=document.createElement("h3");a.textContent=`${e} (${r})`,o.appendChild(a),n[e].forEach((t=>{const n=document.createElement("div");n.className="sub-box",n.textContent=t,n.addEventListener("click",(async()=>{await async function(e,t){try{await Word.run((async n=>{const o=n.document.getSelection();if(o.load("text"),await n.sync(),o.text.includes(e)){const a=o.search(e,{matchCase:!1,matchWholeWord:!0});a.load("items"),await n.sync(),a.items.forEach((e=>{e.insertText(t,"replace"),e.font.highlightColor="white"})),await n.sync(),console.log(`Replaced '${e}' with '${t}' and removed highlight in the document.`)}else console.log(`Selected text does not contain '${e}'.`)}))}catch(e){console.error("Error replacing term:",e)}}(e,t),o.remove()})),o.appendChild(n)})),t.appendChild(o)}))}}(o),await e.sync()}else console.log("No text is selected.")}))}catch(e){console.error("Error displaying glossary:",e)}}()}async function L(){try{await Word.run((async e=>{document.getElementById("app-body").innerHTML='\n      <div id="button-container">\n    \n              <div class="loader" id="loader"></div>\n    \n            <div id="highlighted-text"></div>';const t=e.document.body,n=c.map((e=>{const n=t.search(e.ClinicalTerm,{matchCase:!1,matchWholeWord:!0});return n.load("items"),n}));await e.sync(),n.forEach((e=>{e.items.forEach((e=>{e.font.highlightColor="white"}))})),document.getElementById("app-body").innerHTML='\n      <div id="button-container">\n        <button class="btn btn-secondary me-2 mark-glossary btn-sm" id="applyglossary">Apply Glossary</button>\n      </div>\n',await e.sync(),s=!1,document.getElementById("applyglossary").addEventListener("click",v)})),console.log("Glossary cleared successfully")}catch(e){console.error("Error clearing glossary:",e)}}async function T(){if(!d){s&&await L();const e='\n      <div class="container mt-3">\n        <div class="card">\n          <div class="card-header">\n            <h5 class="card-title">Search Suggestions</h5>\n          </div>\n          <div class="card-body">\n            <div class="form-group">\n              <input type="text" id="search-box" class="form-control" placeholder="Search Suggestions..." autocomplete="off" />\n            </div>\n            <ul id="suggestion-list" class="list-group mt-2"></ul>\n          </div>\n        </div>\n      </div>\n    ';document.getElementById("app-body").innerHTML=e;const t=document.getElementById("search-box"),n=document.getElementById("suggestion-list");function o(){const e=t.value.toLowerCase();n.innerHTML="",a.filter((t=>t.DisplayName.toLowerCase().includes(e))).forEach((e=>{const t=document.createElement("li");t.className="list-group-item list-group-item-action",t.textContent=e.DisplayName,t.onclick=()=>{!async function(e,t){Word.run((async n=>{try{const o=n.document.getSelection();if(await n.sync(),!o)throw new Error("Selection is invalid or not found.");if("TABLE"===t){const t=(new DOMParser).parseFromString(e,"text/html").querySelector("table");if(!t)throw new Error("No table found in the provided HTML.");const a=Array.from(t.querySelectorAll("tr"));if(0===a.length)throw new Error("The table does not contain any rows.");const r=Math.max(...a.map((e=>Array.from(e.querySelectorAll("td, th")).reduce(((e,t)=>e+parseInt(t.getAttribute("colspan")||"1",10)),0)))),s=o.insertParagraph("",Word.InsertLocation.before);if(await n.sync(),!s)throw new Error("Failed to insert the paragraph.");const i=s.insertTable(a.length,r,Word.InsertLocation.after);if(await n.sync(),!i)throw new Error("Failed to insert the table.");const c=new Array(r).fill(0);a.forEach(((e,t)=>{const n=Array.from(e.querySelectorAll("td, th"));let o=0;n.forEach((e=>{for(;c[o]>0;)c[o]--,o++;const n=Array.from(e.childNodes).map((e=>e.nodeType===Node.TEXT_NODE?e.textContent?.trim()||"":e.nodeType===Node.ELEMENT_NODE?e.innerText.trim():"")).filter((e=>e.length>0)).join(" "),a=parseInt(e.getAttribute("colspan")||"1",10),s=parseInt(e.getAttribute("rowspan")||"1",10);o>=r&&(o=r-1);try{i.getCell(t,o).value=n;for(let e=1;e<a;e++)o+e<r&&(i.getCell(t,o+e).value="");if(s>1)for(let e=0;e<a;e++)o+e<r&&(c[o+e]=s-1);o+=a,o>=r&&(o=r-1)}catch(e){console.error("Error setting cell value:",e)}}))}))}else o.insertParagraph(e,Word.InsertLocation.before);await n.sync()}catch(e){console.error("Detailed error:",e)}}))}(e.EditorValue,e.ComponentKeyDataType)},n.appendChild(t)}))}t.addEventListener("input",o)}}function x(e){return e?e.replace(/^"|"$/g,"").replace(/\\n/g,"").replace(/\*\*/g,"").replace(/\\r/g,""):""}window.addEventListener("hashchange",(()=>{"#/dashboard"===window.location.hash&&o&&(o=!1,async function(){try{const o=await fetch(`https://plsdevapp.azurewebsites.net/api/report/id/${t}`,{method:"GET",headers:{Authorization:`Bearer ${e}`}});if(!o.ok)throw new Error("Network response was not ok.");const r=await o.json();document.getElementById("app-body").innerHTML="",document.getElementById("logo-header").innerHTML='\n        <img  id="main-logo" src="./assets/logo.png" alt="" height="60" class="logo">',document.getElementById("header").innerHTML='\n\n    <div class="d-flex justify-content-around">\n    <button class="btn  btn-dark " id="mention">Suggestions</button>\n            <button class="btn  btn-dark " id="aitag">AI Text Panel</button>\n\n        <button class="btn  btn-dark " id="glossary">Glossary</button>\n</div>\n\n',document.getElementById("mention").addEventListener("click",T),document.getElementById("glossary").addEventListener("click",w),document.getElementById("aitag").addEventListener("click",g),l=r.Data;const s=r.Data.Group.find((e=>"AIGroup"===e.DisplayName));i=s?s.Name:"",n=s?s.GroupKey:[],a=r.Data.GroupKeyAll.filter((e=>"TABLE"===e.ComponentKeyDataType||"TEXT"===e.ComponentKeyDataType))}catch(e){console.error("Error fetching glossary data:",e)}}())})),Office.onReady((n=>{n.host===Office.HostType.Word&&(document.getElementById("app-body").style.display="flex",document.getElementById("editor"),window.location.hash="#/login",async function(){try{await Word.run((async n=>{const o=n.document.properties.customProperties;o.load("items"),await n.sync();const a=o.items.find((e=>"DocumentID"===e.key));if(!a)return console.log('Custom property "documentID" not found.'),null;t=a.value,async function(){const t=sessionStorage.getItem("token");console.log(t),t?(e=t,window.location.hash="#/dashboard"):p()}()}))}catch(e){console.error("Error retrieving custom property:",e)}}())}))}(),e=r(14385),t=r.n(e),n=new URL(r(58394),r.b),t()(n)}();
//# sourceMappingURL=taskpane.js.map