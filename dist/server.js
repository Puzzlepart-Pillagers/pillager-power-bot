!function(e){var t={};function i(n){if(t[n])return t[n].exports;var o=t[n]={i:n,l:!1,exports:{}};return e[n].call(o.exports,o,o.exports,i),o.l=!0,o.exports}i.m=e,i.c=t,i.d=function(e,t,n){i.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:n})},i.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},i.t=function(e,t){if(1&t&&(e=i(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var n=Object.create(null);if(i.r(n),Object.defineProperty(n,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var o in e)i.d(n,o,function(t){return e[t]}.bind(null,o));return n},i.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return i.d(t,"a",t),t},i.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},i.p="",i(i.s=3)}([function(e,t){e.exports=require("express-msteams-host")},function(e,t){e.exports=require("debug")},function(e,t){e.exports=require("botbuilder-dialogs")},function(e,t,i){e.exports=i(4)},function(e,t,i){"use strict";Object.defineProperty(t,"__esModule",{value:!0});const n=i(5),o=i(6),r=i(7),s=i(8),a=i(0),c=i(1)("msteams");c("Initializing Microsoft Teams Express hosted App..."),i(9).config();const u=i(10),l=n(),d=process.env.port||process.env.PORT||3007;l.use(n.json({verify:(e,t,i,n)=>{e.rawBody=i.toString()}})),l.use(n.urlencoded({extended:!0})),l.set("views",r.join(__dirname,"/")),l.use(s("tiny")),l.use("/scripts",n.static(r.join(__dirname,"web/scripts"))),l.use("/assets",n.static(r.join(__dirname,"web/assets"))),l.use(a.MsTeamsApiRouter(u)),l.use(a.MsTeamsPageRouter({root:r.join(__dirname,"web/"),components:u})),l.use("/",n.static(r.join(__dirname,"web/"),{index:"index.html"})),l.set("port",d),o.createServer(l).listen(d,()=>{c(`Server running on ${d}`)})},function(e,t){e.exports=require("express")},function(e,t){e.exports=require("http")},function(e,t){e.exports=require("path")},function(e,t){e.exports=require("morgan")},function(e,t){e.exports=require("dotenv")},function(e,t,i){"use strict";Object.defineProperty(t,"__esModule",{value:!0}),t.nonce={},function(e){for(var i in e)t.hasOwnProperty(i)||(t[i]=e[i])}(i(11))},function(e,t,i){"use strict";var n=this&&this.__decorate||function(e,t,i,n){var o,r=arguments.length,s=r<3?t:null===n?n=Object.getOwnPropertyDescriptor(t,i):n;if("object"==typeof Reflect&&"function"==typeof Reflect.decorate)s=Reflect.decorate(e,t,i,n);else for(var a=e.length-1;a>=0;a--)(o=e[a])&&(s=(r<3?o(s):r>3?o(t,i,s):o(t,i))||s);return r>3&&s&&Object.defineProperty(t,i,s),s},o=this&&this.__awaiter||function(e,t,i,n){return new(i||(i=Promise))((function(o,r){function s(e){try{c(n.next(e))}catch(e){r(e)}}function a(e){try{c(n.throw(e))}catch(e){r(e)}}function c(e){e.done?o(e.value):new i((function(t){t(e.value)})).then(s,a)}c((n=n.apply(e,t||[])).next())}))};Object.defineProperty(t,"__esModule",{value:!0});const r=i(0),s=i(1),a=i(2),c=i(12),u=i(13),l=i(14),d=i(16);s("msteams");let p=class{constructor(e){this.activityProc=new d.TeamsActivityProcessor,this.conversationState=e,this.dialogState=e.createProperty("dialogState"),this.dialogs=new a.DialogSet(this.dialogState),this.dialogs.add(new u.default("help")),console.log("",process.env.MICROSOFT_APP_ID),console.log("",process.env.MICROSOFT_APP_PASSWORD),this.activityProc.messageActivityHandler={onMessage:e=>o(this,void 0,void 0,(function*(){const t=d.TeamsContext.from(e);switch(e.activity.type){case c.ActivityTypes.Message:const i=t?t.getActivityTextWithoutMentions().toLowerCase():e.activity.text;if(i.startsWith("hello"))return void(yield e.sendActivity("Oh, hello to you as well!"));if(i.startsWith("help")){const t=yield this.dialogs.createContext(e);yield t.beginDialog("help")}else yield e.sendActivity("I'm terribly sorry, but my master hasn't trained me to do anything yet...")}return this.conversationState.saveChanges(e)}))},this.activityProc.conversationUpdateActivityHandler={onConversationUpdateActivity:e=>o(this,void 0,void 0,(function*(){if(e.activity.membersAdded&&0!==e.activity.membersAdded.length)for(const t in e.activity.membersAdded)if(e.activity.membersAdded[t].id===e.activity.recipient.id){const t=c.CardFactory.adaptiveCard(l.default);yield e.sendActivity({attachments:[t]})}}))},this.activityProc.messageReactionActivityHandler={onMessageReaction:e=>o(this,void 0,void 0,(function*(){const t=e.activity.reactionsAdded;t&&t[0]&&(yield e.sendActivity({textFormat:"xml",text:`That was an interesting reaction (<b>${t[0].type}</b>)`}))}))}}onTurn(e){return o(this,void 0,void 0,(function*(){yield this.activityProc.processIncomingActivity(e)}))}};p=n([r.BotDeclaration("/api/messages",new c.MemoryStorage,process.env.MICROSOFT_APP_ID,process.env.MICROSOFT_APP_PASSWORD)],p),t.PowerPillager=p},function(e,t){e.exports=require("botbuilder")},function(e,t,i){"use strict";var n=this&&this.__awaiter||function(e,t,i,n){return new(i||(i=Promise))((function(o,r){function s(e){try{c(n.next(e))}catch(e){r(e)}}function a(e){try{c(n.throw(e))}catch(e){r(e)}}function c(e){e.done?o(e.value):new i((function(t){t(e.value)})).then(s,a)}c((n=n.apply(e,t||[])).next())}))};Object.defineProperty(t,"__esModule",{value:!0});const o=i(2);class r extends o.Dialog{constructor(e){super(e)}beginDialog(e,t){return n(this,void 0,void 0,(function*(){return e.context.sendActivity("I'm just a friendly but rather stupid bot, and right now I don't have any valuable help for you!"),yield e.endDialog()}))}}t.default=r},function(e,t,i){"use strict";Object.defineProperty(t,"__esModule",{value:!0});const n=i(15);t.default=n},function(e){e.exports={$schema:"http://adaptivecards.io/schemas/adaptive-card.json",type:"AdaptiveCard",version:"1.0",body:[{type:"Image",url:"https://pillagers-teams-bot.herokuapp.com/assets/icon.png",size:"stretch"},{type:"TextBlock",spacing:"medium",size:"default",weight:"bolder",text:"Welcome to Pillager bot",wrap:!0,maxLines:0},{type:"TextBlock",size:"default",isSubtle:!0,text:"Hello, nice to meet you!",wrap:!0,maxLines:0}],actions:[{type:"Action.OpenUrl",title:"Learn more about Yo Teams",url:"https://aka.ms/yoteams"},{type:"Action.OpenUrl",title:"Pillager bot",url:"https://pillagers-teams-bot.herokuapp.com"}]}},function(e,t){e.exports=require("botbuilder-teams")}]);
//# sourceMappingURL=server.js.map