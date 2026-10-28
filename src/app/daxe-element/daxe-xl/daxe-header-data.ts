export const fieldData = [
  {
    "fieldId": "status",
    "label": "Status",
    "textColor": "#000000",
    "group": "Status",
    "column": "A",
    "columnNumber": 0,
    "bold": true,
    "width": 150,
    "bgColor": "#C0C0C0"
  },
  {
    "fieldId": "success_error",
    "label": "Error Message",
    "textColor": "#000000",
    "group": "Status",
    "column": "B",
    "columnNumber": 1,
    "bold": true,
    "width": 150
  },
  {
    "fieldId": "mdo_number",
    "label": "MDO Number",
    "textColor": "#ff0000",
    "group": "Record Number",
    "column": "C",
    "columnNumber": 2,
    "bold": true,
    "width": 150,
    "bgColor": "#C0C0C0"
  },
  {
    "column": "D",
    "columnNumber": 3,
    "fieldId": "FLD_060471536",
    "label": "SAP Material Number",
    "textColor": "#000000",
    "bold": true,
    "width": 150,
    "bgColor": "#C0C0C0",
    "group": "Header data",
    "groupBgColor": "#FFFF00",
    "structureId": 1,
    "isGridField": false,
    "isKeyField": false,
    "isStaticClassCharacteristics": false,
    "isStaticClassCharacteristicsUoM": false,
    "isClassificationField": false,
    "fieldType": "FIELD",
    "classCtrl": false
  },
  {
    "column": "E",
    "columnNumber": 4,
    "fieldId": "FLD_657134110",
    "label": "Material Type",
    "textColor": "#000000",
    "bold": true,
    "width": 150,
    "bgColor": "#C0C0C0",
    "group": "Header data",
    "groupBgColor": "#FFFF00",
    "structureId": 1,
    "isGridField": false,
    "isKeyField": false,
    "isStaticClassCharacteristics": false,
    "isStaticClassCharacteristicsUoM": false,
    "isClassificationField": false,
    "fieldType": "FIELD",
    "classCtrl": false
  },
  {
    "column": "F",
    "columnNumber": 5,
    "fieldId": "SHORT_DESCRIPTION",
    "label": "Short description",
    "textColor": "#000000",
    "bold": true,
    "width": 150,
    "group": "Descriptions",
    "groupBgColor": "#808080"
  },
  {
    "column": "G",
    "columnNumber": 6,
    "fieldId": "LONG_DESCRIPTION",
    "label": "Long description",
    "textColor": "#000000",
    "bold": true,
    "width": 150,
    "group": "Descriptions",
    "groupBgColor": "#808080"
  },
  {
    "column": "H",
    "columnNumber": 7,
    "fieldId": "CLASSIFICATION_CLASS",
    "label": "Class",
    "textColor": "#ff0000",
    "bold": true,
    "width": 150,
    "group": "Descriptions",
    "groupBgColor": "#808080",
    "isClassificationField": false,
    "fieldType": "Static",
    "metadata": [
      {
        "fieldId": "CLASSIFICATION_CLASS",
        "description": "Class",
        "pickList": "1",
        "structureId": 1,
        "isGridColumn": false
      }
    ]
  }
];

export const fieldDataImport = [...fieldData];

/** Dynamic CLASS_TYPE data is stored independently and joined to the primary
 * record by Record Number. The host recognises Characteristic/Value metadata
 * and provides the appropriate attribute-value editor. */
export const dynamicClassificationSheet = undefined;

export const extentionHeaderFields = fieldData.map(field => field.fieldId);

export const groupBgColors: Record<string, string> = {
  "Status": "#b1e39d",
  "Record Number": "#f7c7ac",
  "Header data": "#FFFF00",
  "Descriptions": "#808080"
};

export const controlData = {
  "moduleId": "968371",
  "massId": "",
  "layoutId": "53d767b2-5443-4d9b-9d63-f82ce6664a8f",
  "tenantId": "977625",
  "userId": "",
  "validationRequired": true
};

export const codeFields = [
  "FLD_060471536",
  "FLD_657134110"
];

export const data = {
  controlData,
  flowId: "Update_description_Flow.Updatedescription",
  formId: "53d767b2-5443-4d9b-9d63-f82ce6664a8f",
  formDesc: "Update description",
  stepId: "_2938A7F6-1695-4740-956D-6525FAE8EC5C",
  schemaId: "",
  childFormId: "",
  classificationEnabled: true,
  generationProfile: {
  "process": "create",
  "autoImportBeforeValidation": false,
  "changeImportModes": [],
  "acceptsHostImportResponse": false,
  "descriptionGeneration": true,
  "classificationEnabled": true,
  "staticClassificationAttributes": false,
  "dynamicClassificationAttributes": false
},
  businessRules: [] as any[],
  baseStructure: [{
    moduleId: controlData.moduleId,
    formId: "53d767b2-5443-4d9b-9d63-f82ce6664a8f",
    isHeader: true,
    isGrid: false,
    isHierarchy: false
  }] as any[]
};

export const TRANSLATIONS = {
  en: {
    'Generate Description': 'Generate Description',
    'Apply Data Rules': 'Apply Data Rules',
    'Validate Records': 'Validate Records',
    'Submit Records': 'Submit Records',
    'Send for Correction': 'Send for Correction',
    'Loading data...': 'Loading data...',
    'Status': 'Status',
    'Error Message': 'Error Message',
    'MDO Number': 'MDO Number'
  }
} as { [lang: string]: { [key: string]: string } };
