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
    // Sort dropDownValues alphabetically by label in ascending order
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
      ...Object.keys(excel?.config[uid]?.relation),
    ];

    let where = {};

    if (excel?.config[uid]?.locale == "true") {
      where = {
        locale: "en",
      };
    }

    let count = await strapi.db.query(uid).count(where);

    let tableData = await this.restructureData(response, excel?.config[uid]);

    // Sort dropDownValues alphabetically by label in ascending order

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
  
      // Dynamically determine column headers
      let headers = [
        ...excel?.config[uid]?.columns,
        ...[].concat(...Object.keys(excel?.config[uid]?.relation).map(key => {
          const relation = excel?.config[uid]?.relation[key];
          const columns = relation.columns || [relation.column];  // Handle both 'columns' and 'column'
          return columns.map(column => `${key}_${column}`);
        }))
      ];
  
      // Transform headers to human-readable format
      let headerRestructure = headers.map(header => {
        return header.split("_").map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(" ");
      });
  
      // Define worksheet columns with headers
      worksheet.columns = headers.map((header, index) => ({
        header: headerRestructure[index],
        key: header,
        width: 20,
      }));
  
      // Add data to the worksheet
      excelData.forEach(row => {
        worksheet.addRow(row);
      });
  
      // Additional formatting
      worksheet.columns.forEach(column => {
        column.alignment = { wrapText: true };
      });
      worksheet.views = [{ state: "frozen", xSplit: 0, ySplit: 1, topLeftCell: "A" }];
  
      const buffer = await workbook.xlsx.writeBuffer();
      return buffer;
    } catch (error) {
      console.error("Error writing buffer:", error);
    }
  },
  
  async restructureObject(inputObject, uid, limit, offset) {
    let excel = strapi.config.get("excel");
  
    let where = {};
    if (excel?.config[uid]?.locale == "true") {
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
        select: inputObject.relation[key].columns || inputObject.relation[key].column,  // Ensure both 'columns' and 'column' are handled
      };
    }
  
    return restructuredObject;
  },
  
  async restructureData(data, objectStructure) {
    return data.map((item) => {
      const restructuredItem = {};
  
      // Handle main data columns
      objectStructure.columns.forEach((key) => {
        if (key in item) {
          restructuredItem[key] = item[key];
        }
      });
  
      // Handle relations
      for (const key in objectStructure.relation) {
        const relation = objectStructure.relation[key];
        const columns = relation.columns || [relation.column]; // Ensure both 'columns' and 'column' are handled as arrays
        if (item[key] && typeof item[key] === "object") {
          columns.forEach(column => {
            const value = Array.isArray(item[key]) ? item[key].map(obj => obj[column] || '').join(", ") : item[key][column] || '';
            restructuredItem[`${key}_${column}`] = value; // Create a unique key for each column in the relation
          });
        } else {
          columns.forEach(column => {
            restructuredItem[`${key}_${column}`] = ''; // Set empty if no data
          });
        }
      }
  
      return restructuredItem;
    });
  }
  
});
