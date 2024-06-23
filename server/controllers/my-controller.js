"use strict";
const ExcelJS = require("exceljs");

module.exports = ({ strapi }) => ({
  async getDropDownData() {
    let excel = strapi.config.get("excel");
    let dropDownValues = [];
    let array = Object.keys(excel?.config);

    strapi?.db?.config?.models?.forEach((element) => {
      if (element?.kind == "collectionType") {
        array?.forEach((data) => {
          if (element?.uid?.startsWith(data)) {
            dropDownValues.push({
              label: element?.info?.displayName,
              value: element?.uid,
            });
          }
        });
      }
    });
    dropDownValues.sort((a, b) => a.label.localeCompare(b.label));
    return {
      data: dropDownValues,
    };
  },

  async getTableData(ctx) {
    let excel = strapi.config.get("excel");
    let uid = ctx?.query?.uid;
    let limit = ctx?.query?.limit;
    let offset = ctx?.query?.offset;
    let query = await this.restructureObject(
      excel?.config[uid],
      uid,
      limit,
      offset
    );

    let response = await strapi.db.query(uid).findMany(query);

    let header = [
      ...excel?.config[uid]?.columns,
      ...Object.keys(excel?.config[uid]?.relation).map(rKey => excel.config[uid].relation[rKey].columns).flat(),
    ];

    let where = {};
    if (excel?.config[uid]?.locale === "true") {
      where = {
        locale: "en",
      };
    }

    let count = await strapi.db.query(uid).count(where);

    let tableData = await this.restructureData(response, excel?.config[uid]);

    return {
      data: tableData,
      count: count,
      columns: header,
    };
  },

  async downloadExcel(ctx) {
    try {
      let excel = strapi.config.get("excel");
      let uid = ctx?.query?.uid;
      let query = await this.restructureObject(excel?.config[uid], uid);
      let response = await strapi.db.query(uid).findMany(query);
      let excelData = await this.restructureData(response, excel?.config[uid]);
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet 1");

      let headers = [
        ...excel?.config[uid]?.columns,
        ...Object.keys(excel?.config[uid]?.relation).map(rKey => excel.config[uid].relation[rKey].columns).flat(),
      ];

      let headerRestructure = headers.map(element =>
        element.split("_").map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(" ")
      );

      worksheet.columns = headers.map((header, index) => ({
        header: headerRestructure[index],
        key: header,
        width: 20,
      }));

      excelData?.forEach(row => {
        worksheet.addRow(row);
      });

      worksheet.columns.forEach(column => {
        column.alignment = { wrapText: true };
      });

      worksheet.views = [
        { state: "frozen", xSplit: 0, ySplit: 1, topLeftCell: "A2" },
      ];

      const buffer = await workbook.xlsx.writeBuffer();
      return buffer;
    } catch (error) {
      console.error("Error writing buffer:", error);
    }
  },

  async restructureObject(inputObject, uid, limit, offset) {
    let excel = strapi.config.get("excel");

    let where = {};
    if (excel?.config[uid]?.locale === "true") {
      where = {
        locale: "en",
      };
    }
    let orderBy = {
      id: "asc",
    };

    const restructuredObject = {
      select: inputObject.columns || "*",
      populate: {},
      where,
      orderBy,
      limit: limit,
      offset: offset,
    };

    for (const key in inputObject.relation) {
      restructuredObject.populate[key] = {
        select: inputObject.relation[key].columns,
      };
    }

    return restructuredObject;
  },

  async restructureData(data, objectStructure) {
    return data.map((item) => {
      const restructuredItem = {};

      objectStructure.columns.forEach((key) => {
        if (key in item) {
          restructuredItem[key] = item[key];
        }
      });

      for (const key in objectStructure.relation) {
        if (key in item && item[key]) {
          const columns = objectStructure.relation[key].columns;
          if (typeof item[key] === "object" && item[key] !== null) {
            try {
              restructuredItem[key] = columns.map(column => {
                return Array.isArray(item[key])
                  ? item[key].map(obj => obj ? obj[column] : '').join(", ")
                  : item[key][column] || '';
              }).join(" | ");
            } catch (error) {
              console.error(`Error restructuring data for key '${key}':`, error);
            }
          } else {
            restructuredItem[key] = '';
          }
        } else {
          restructuredItem[key] = '';
        }
      }

      return restructuredItem;
    });
  }
});
