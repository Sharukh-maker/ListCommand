"use strict";
var __decorate =
  (this && this.__decorate) ||
  function (decorators, target, key, desc) {
    var c = arguments.length,
      r =
        c < 3
          ? target
          : desc === null
          ? (desc = Object.getOwnPropertyDescriptor(target, key))
          : desc,
      d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function")
      r = Reflect.decorate(decorators, target, key, desc);
    else
      for (var i = decorators.length - 1; i >= 0; i--)
        if ((d = decorators[i]))
          r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
  };
var __awaiter =
  (this && this.__awaiter) ||
  function (thisArg, _arguments, P, generator) {
    function adopt(value) {
      return value instanceof P
        ? value
        : new P(function (resolve) {
            resolve(value);
          });
    }
    return new (P || (P = Promise))(function (resolve, reject) {
      function fulfilled(value) {
        try {
          step(generator.next(value));
        } catch (e) {
          reject(e);
        }
      }
      function rejected(value) {
        try {
          step(generator["throw"](value));
        } catch (e) {
          reject(e);
        }
      }
      function step(result) {
        result.done
          ? resolve(result.value)
          : adopt(result.value).then(fulfilled, rejected);
      }
      step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
  };
Object.defineProperty(exports, "__esModule", { value: true });
const decorators_1 = require("@microsoft/decorators");
const sp_listview_extensibility_1 = require("@microsoft/sp-listview-extensibility");
const xlsx = require("xlsx");
const sp_http_1 = require("@microsoft/sp-http");
const LOG_SOURCE = "ExportListItemsCommandSet";
class ExportListItemsCommandSet extends sp_listview_extensibility_1.BaseListViewCommandSet {
  onInit() {
    this.Initiate();
    return Promise.resolve();
  }
  onListViewUpdated(event) {
    const exportCommand = this.tryGetCommand("COMMAND_1");
    if (exportCommand) {
      // This command should be hidden unless exactly one row is selected.
      exportCommand.visible = event.selectedRows.length > 0;
    }
  }
  onExecute(event) {
    let _grid;
    // One dirty fix for LinkTitle column internal name
    let index = this._viewColumns.indexOf("LinkTitle");
    if (index !== -1) {
      this._viewColumns[index] = "Title";
    }
    switch (event.itemId) {
      case "COMMAND_1":
        if (event.selectedRows.length > 0) {
          _grid = new Array(event.selectedRows.length);
          _grid[0] = this._viewColumns;
          event.selectedRows.forEach((row, index) => {
            let _row = [],
              i = 0;
            this._viewColumns.forEach((viewColumn) => {
              _row[i++] = this._getFieldValueAsText(
                row.getValueByName(viewColumn)
              );
            });
            _grid[index + 1] = _row;
          });
        }
        break;
      default:
        throw new Error("Unknown command");
    }
    this.writeToExcel(_grid);
  }
  /*
    Some brute force to identify the type of field and return the text value of the field, trying to avoid one more rest call for field types
    Tested, Single line, Multiline, Choice, Number, Boolean, Lookup and Managed metadata,
    */
  _getFieldValueAsText(field) {
    let fieldValue;
    switch (typeof field) {
      case "object": {
        if (field instanceof Array) {
          if (!field.length) {
            fieldValue = "";
          }
          // people
          else if (field[0].title) {
            fieldValue = field.map((value) => value.title).join();
          }
          // lookup
          else if (field[0].lookupValue) {
            fieldValue = field.map((value) => value.lookupValue).join();
          }
          // managed metadata
          else if (field[0].Label) {
            fieldValue = field.map((value) => value.Label).join();
          }
          // choice and others
          else {
            fieldValue = field.join();
          }
        }
        break;
      }
      default: {
        fieldValue = field;
      }
    }
    return fieldValue;
  }
  writeToExcel(data) {
    let ws = xlsx.utils.aoa_to_sheet(data);
    let wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "selected-items");
    xlsx.writeFile(wb, `${this._listTitle}.xlsx`);
  }
  getViewColumns() {
    return __awaiter(this, void 0, void 0, function* () {
      const currentWebUrl = this.context.pageContext.web.absoluteUrl;
      this._listTitle = this.context.pageContext.legacyPageContext.listTitle;
      const viewId = this.context.pageContext.legacyPageContext.viewId
        .replace("{", "")
        .replace("}", "");
      this.context.spHttpClient
        .get(
          `${currentWebUrl}/_api/lists/getbytitle('${this._listTitle}')/Views('${viewId}')/ViewFields`,
          sp_http_1.SPHttpClient.configurations.v1
        )
        .then((res) => {
          res.json().then((viewColumnsResponse) => {
            this._viewColumns = viewColumnsResponse.Items;
          });
        });
    });
  }
  Initiate() {
    return __awaiter(this, void 0, void 0, function* () {
      yield this.getViewColumns();
    });
  }
}
__decorate(
  [decorators_1.override],
  ExportListItemsCommandSet.prototype,
  "onInit",
  null
);
__decorate(
  [decorators_1.override],
  ExportListItemsCommandSet.prototype,
  "onListViewUpdated",
  null
);
__decorate(
  [decorators_1.override],
  ExportListItemsCommandSet.prototype,
  "onExecute",
  null
);
exports.default = ExportListItemsCommandSet;
//# sourceMappingURL=ExportItemsCommandSet.js.map
