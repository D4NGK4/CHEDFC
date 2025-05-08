function getATemplate() {
  const result = DocumentGenerator.getTemplate("Test Template");
  Logger.log(JSON.stringify(result, null, 2));
}

function getAllTemplates() {
  try {
    const response = DocumentGenerator.getTemplates();

    if (!response || !response.header || !response.data) {
      Logger.log("Invalid templates response structure");
      return {};
    }

    const allTemplates = {};
    const headerIndex = {};

    // Create a map of header names to their indices
    response.header.forEach((header, index) => {
      headerIndex[header] = index;
    });

    // Process each template row
    response.data.forEach((row) => {
      const fileName = row[headerIndex["File Name"]];

      if (!fileName) {
        Logger.log("Skipping row with missing File Name");
        return;
      }

      allTemplates[fileName] = {
        "File Name": fileName,
        "File ID": row[headerIndex["File ID"]],
        Fields: row[headerIndex["Fields"]],
        "Fields for Filename": row[headerIndex["Fields for Filename"]],
        isPrivate: row[headerIndex["isPrivate"]] === "TRUE",
        isTrashed: row[headerIndex["isTrashed"]] === "TRUE",
        Email: row[headerIndex["Email"]],
        Timestamp: row[headerIndex["Timestamp"]],
      };
    });

    Logger.log("Successfully loaded all templates data");
    Logger.log(allTemplates);
    return allTemplates;
  } catch (error) {
    Logger.log(`Error loading templates: ${error}`);
    return {};
  }
}

function mapTemplateData(templateName) {
  const allTemplates = this.getAllTemplates();

  if (!allTemplates[templateName]) {
    Logger.log(`Template "${templateName}" not found in loaded templates`);
    return {};
  }

  Logger.log(`Mapped Template Data for ${templateName}`);
  return allTemplates[templateName];
}

function initializeFormRequest(templateId) {
  try {
    const templateData = this.mapTemplateData(templateId);
    if (!templateData || !templateData["File Name"]) {
      throw new Error("No template data found");
    }

    const tempControlNumber = "TEMP-" + Utilities.getUuid().substring(0, 8);

    return {
      status: "success",
      controlNumber: tempControlNumber,
      formData: {
        Template_ID: templateData["File Name"],
        Year: new Date().getFullYear(),
        Control_Number: tempControlNumber,
        Form_Fields: templateData["Fields"],
        File_Name: templateData["File Name"],
        File_ID: templateData["File ID"],
        Is_Temporary: true,
      },
      templateData: templateData,
    };
  } catch (error) {
    Logger.log("Error in initializeFormRequest: " + error.toString());
    throw error;
  }
}
