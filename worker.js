import { Subject } from 'threads/observable'
import { expose } from 'threads/worker'

const SocialCalc = require('socialcalc/dist/SocialCalc')

var window = this;

var Node
Node = (function () {
  Node.displayName = 'Node'
  var prototype = Node.prototype, constructor = Node

  function Node (tag, attrs, style, elems, raw) {
    this.tag = tag != null ? tag : 'div'
    this.attrs = attrs != null
      ? attrs
      : {}
    this.style = style != null
      ? style
      : {}
    this.elems = elems != null
      ? elems
      : []
    this.raw = raw != null ? raw : ''
  }

  Object.defineProperty(Node.prototype, 'id', {
    set: function (id) {
      this.attrs.id = id
    },
    configurable: true,
    enumerable: true
  })
  Object.defineProperty(Node.prototype, 'width', {
    set: function (width) {
      this.attrs.width = width
    },
    configurable: true,
    enumerable: true
  })
  Object.defineProperty(Node.prototype, 'height', {
    set: function (height) {
      this.attrs.height = height
    },
    configurable: true,
    enumerable: true
  })
  Object.defineProperty(Node.prototype, 'className', {
    set: function ($class) {
      this.attrs['class'] = $class
    },
    configurable: true,
    enumerable: true
  })
  Object.defineProperty(Node.prototype, 'colSpan', {
    set: function (colspan) {
      this.attrs.colspan = colspan
    },
    configurable: true,
    enumerable: true
  })
  Object.defineProperty(Node.prototype, 'rowSpan', {
    set: function (rowspan) {
      this.attrs.rowspan = rowspan
    },
    configurable: true,
    enumerable: true
  })
  Object.defineProperty(Node.prototype, 'title', {
    set: function (title) {
      this.attrs.title = title
    },
    configurable: true,
    enumerable: true
  })
  Object.defineProperty(Node.prototype, 'innerHTML', {
    set: function (raw) {
      this.raw = raw
    },
    get: function () {
      var e
      return this.raw || (function () {
        var i$, ref$, len$, results$ = []
        for (i$ = 0, len$ = (ref$ = this.elems).length; i$ < len$; ++i$) {
          e = ref$[i$]
          results$.push(e.outerHTML)
        }
        return results$
      }.call(this)).join('\n')
    },
    configurable: true,
    enumerable: true
  })
  Object.defineProperty(Node.prototype, 'outerHTML', {
    get: function () {
      var tag, attrs, style, css, k, v
      tag = this.tag, attrs = this.attrs, style = this.style
      css = style.cssText || (function () {
        var ref$, results$ = []
        for (k in ref$ = style) {
          v = ref$[k]
          results$.push(k.replace(/[A-Z]/g, '-$&').toLowerCase() + ':' + v)
        }
        return results$
      }()).join(';')
      if (css) {
        attrs.style = css
      } else {
        delete attrs.style
      }
      return '<' + tag + (function () {
        var ref$, results$ = []
        for (k in ref$ = attrs) {
          v = ref$[k]
          results$.push(' ' + k + '="' + v + '"')
        }
        return results$
      }()).join('') + '>' + this.innerHTML + '</' + tag + '>'
    },
    configurable: true,
    enumerable: true
  })
  Node.prototype.appendChild = function (it) {
    return this.elems.push(it)
  }
  return Node
}())

SocialCalc.document == null && (SocialCalc.document = {})
SocialCalc.document.createElement = function (it) {
  return new Node(it)
}

function execute (ref$) {
  var type, ref, snapshot, command, room, log, ref1$, commandParameters, csv, alert, ss, parts, cmdstr, line;
  type = ref$.type, ref = ref$.ref, snapshot = ref$.snapshot, command = ref$.command, room = ref$.room, log = (ref1$ = ref$.log) != null
    ? ref1$
    : [];

  switch (type) {
    case 'cmd':
      commandParameters = command.split(" ");

      if (commandParameters[0] === 'settimetrigger') {
        return {
          type: 'setcrontrigger',
          timetriggerdata: {
            cell: commandParameters[1],
            times: commandParameters[2]
          }
        }
      }

      if (commandParameters[0] === 'sendemail') {
        return {
          type: 'sendemailout',
          emaildata: {
            to: commandParameters[1].replace(/%20/g, ' '),
            subject: commandParameters[2].replace(/%20/g, ' '),
            body: commandParameters[3].replace(/%20/g, ' ')
          }
        }
      }

      return window.ss.ExecuteCommand(command);

    case 'recalc':
      return SocialCalc.RecalcLoadedSheet(ref, snapshot, true);

    case 'clearCache':
      return SocialCalc.Formula.SheetCache.sheets = {};

    case 'exportSave':
      return {
        type: 'save',
        save: window.ss.CreateSheetSave()
      };

    case 'exportHTML':
      return {
        type: 'html',
        html: window.ss.CreateSheetHTML()
      };

    case 'exportCSV':
      csv = window.ss.SocialCalc.ConvertSaveToOtherFormat(window.ss.CreateSheetSave(), 'csv');
      return {
        type: 'csv',
        csv: csv
      };

    case 'exportCells':
      return {
        type: 'cells',
        cells: window.ss.cells
      };

    case 'init':
      const subject = new Subject()

      SocialCalc.CreateAuditString = function(){
        return "";
      };

      SocialCalc.CalculateEditorPositions = function(){};
      SocialCalc.Popup.Types.List.Create = function(){};
      SocialCalc.Popup.Types.ColorChooser.Create = function(){};
      SocialCalc.Popup.Initialize = function(){};
      SocialCalc.RecalcInfo.LoadSheet = function(ref){
        if (/[^.=_a-zA-Z0-9]/.exec(ref)) {
          return;
        }

        ref = ref.toLowerCase();
        subject.next({
          type: 'load-sheet',
          ref: ref
        });

        return true;
      };

      window.setTimeout = function(cb, ms){
        return process.nextTick(cb);
      };
      window.clearTimeout = function(){};
      window.alert = alert = function(){};

      window.ss = ss = new SocialCalc.SpreadsheetControl;
      ss.SocialCalc = SocialCalc;
      ss._room = room;

      if (snapshot) {
        parts = ss.DecodeSpreadsheetSave(snapshot);
      }

      ss.editor.StatusCallback.EtherCalc = {
        func: function(editor, status, arg){
          var newSnapshot;
          if (status !== 'doneposcalc') {
            return;
          }
          newSnapshot = ss.CreateSpreadsheetSave();
          if (ss._snapshot === newSnapshot) {
            return;
          }
          ss._snapshot = newSnapshot;
          return subject.next({
            type: 'snapshot',
            snapshot: newSnapshot
          });
        }
      };

      if (parts != null) {
        if (parts.sheet) {
          ss.sheet.ResetSheet();
          ss.ParseSheetSave(snapshot.substring(parts.sheet.start, parts.sheet.end));
        }
        if (parts.edit) {
          ss.editor.LoadEditorSettings(snapshot.substring(parts.edit.start, parts.edit.end));
        }
      }

      cmdstr = (function(){
        var i$, ref$, len$, results$ = [];
        for (i$ = 0, len$ = (ref$ = log).length; i$ < len$; ++i$) {
          line = ref$[i$];
          if (!/^re(calc|display)$/.test(line)) {
            results$.push(line);
          }
        }
        return results$;
      }()).join("\n");

      if (cmdstr.length) {
        cmdstr += "\n";
      }

      ss.context.sheetobj.ScheduleSheetCommands("set sheet defaulttextvalueformat text-wiki\n" + cmdstr + "recalc\n", false, true);

      return subject
  }
}

function eval_ (ref$) {
  var snapshot, log, ref1$, code, parts, save, ss, cmdstr, line, e;
  snapshot = ref$.snapshot, log = (ref1$ = ref$.log) != null
    ? ref1$
    : [], code = ref$.code;

  try {
    parts = SocialCalc.SpreadsheetControlDecodeSpreadsheetSave("", snapshot);
    save = snapshot.substring(parts.sheet.start, parts.sheet.end);

    window.clearTimeout = function(){};
    window.ss = ss = new SocialCalc.SpreadsheetControl;
    ss.sheet.ResetSheet();
    ss.ParseSheetSave(save);
    if (log != null && log.length) {
      cmdstr = (function(){
        var i$, ref$, len$, results$ = [];
        for (i$ = 0, len$ = (ref$ = log).length; i$ < len$; ++i$) {
          line = ref$[i$];
          if (!/^re(calc|display)$/.test(line) && line !== "set sheet defaulttextvalueformat text-wiki") {
            results$.push(line);
          }
        }
        return results$;
      }()).join("\n");
      if (cmdstr.length) {
        cmdstr += "\n";
      }

      return new Promise(res => {
        ss.editor.StatusCallback.EtherCalc = {
          func: function (editor, status, arg) {
            if (status !== 'doneposcalc') {
              return res();
            }

            return res(eval(code));
          }
        };

        ss.context.sheetobj.ScheduleSheetCommands(cmdstr, false, true);
      })
    } else {
      return eval(code);
    }
  } catch (e) {
    return "ERROR: " + e;
  }
}

function exportCSV (ref$) {
  var snapshot, log, ref1$, parts, save, cmdstr, line, ss, e;
  snapshot = ref$.snapshot, log = (ref1$ = ref$.log) != null
    ? ref1$
    : [];

  try {
    parts = SocialCalc.SpreadsheetControlDecodeSpreadsheetSave("", snapshot);
    save = snapshot.substring(parts.sheet.start, parts.sheet.end);
    if (log != null && log.length) {
      cmdstr = (function(){
        var i$, ref$, len$, results$ = [];
        for (i$ = 0, len$ = (ref$ = log).length; i$ < len$; ++i$) {
          line = ref$[i$];
          if (!/^re(calc|display)$/.test(line) && line !== "set sheet defaulttextvalueformat text-wiki") {
            results$.push(line);
          }
        }
        return results$;
      }()).join("\n");
      if (cmdstr.length) {
        cmdstr += "\n";
      }
      window.setTimeout = function(cb, ms){
        return process.nextTick(cb);
      };
      window.clearTimeout = function(){};
      window.ss = ss = new SocialCalc.SpreadsheetControl;
      ss.sheet.ResetSheet();
      ss.ParseSheetSave(save);

      return new Promise(res => {
        ss.editor.StatusCallback.EtherCalc = {
          func: function (editor, status, arg) {
            var save;
            if (status !== 'doneposcalc') {
              return res();
            }

            save = ss.CreateSheetSave();
            return res(SocialCalc.ConvertSaveToOtherFormat(save, 'csv'));
          }
        };

        ss.context.sheetobj.ScheduleSheetCommands(cmdstr, false, true);
      })
    } else {
      return SocialCalc.ConvertSaveToOtherFormat(save, 'csv');
    }
  } catch (e) {
    return "ERROR: " + e;
  }
}

function exportHTML (ref$) {
  var snapshot, log, ref1$, parts, save, ss, cmdstr, line, e;
  snapshot = ref$.snapshot, log = (ref1$ = ref$.log) != null
    ? ref1$
    : [];
  try {
    parts = SocialCalc.SpreadsheetControlDecodeSpreadsheetSave("", snapshot);
    save = snapshot.substring(parts.sheet.start, parts.sheet.end);
    window.setTimeout = function(cb, ms){
      return process.nextTick(cb);
    };
    window.clearTimeout = function(){};
    window.ss = ss = new SocialCalc.SpreadsheetControl;
    ss.sheet.ResetSheet();
    ss.ParseSheetSave(save);

    if (log != null && log.length) {
      cmdstr = (function(){
        var i$, ref$, len$, results$ = [];
        for (i$ = 0, len$ = (ref$ = log).length; i$ < len$; ++i$) {
          line = ref$[i$];
          if (!/^re(calc|display)$/.test(line) && line !== "set sheet defaulttextvalueformat text-wiki") {
            results$.push(line);
          }
        }
        return results$;
      }()).join("\n");

      if (cmdstr.length) {
        cmdstr += "\n";
      }

      return new Promise(res => {
        ss.editor.StatusCallback.EtherCalc = {
          func: function (editor, status, arg) {
            if (status !== 'doneposcalc') {
              return res();
            }

            return res(ss.CreateSheetHTML());
          }
        };

        ss.context.sheetobj.ScheduleSheetCommands(cmdstr, false, true);
      })
    } else {
      return ss.CreateSheetHTML();
    }
  } catch (e) {
    return "ERROR: " + e;
  }
}

expose({
  execute,
  exportCSV,
  exportHTML,
  eval_
})
