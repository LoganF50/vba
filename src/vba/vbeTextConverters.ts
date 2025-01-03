import {
  getModuleConstants,
  getModuleOptions,
  getPrivateVariables,
  getVBAEnums,
  getVBAProcedures,
} from "./codeMatchers";
import { VBAParam, VBAProcedure } from "./types";

type CommentBlockProps = {
  params?: VBAParam[];
  requirements?: string[];
  returns?: string;
};

export function getCommentBlock({
  params,
  requirements,
  returns,
}: CommentBlockProps) {
  const startLine = "'/**\n";
  const linePrefix = "' * @";
  const endLine = "' */\n";
  let str = "";

  str += linePrefix + `description\n`;

  if (params) {
    params.forEach((p) => {
      const pType = p.isParamArray ? "...Variant" : p.dataType;
      let defValue = "";

      if (typeof p.defaultValue !== "undefined") {
        defValue = `=${p.defaultValue}`;
      }

      let pName = p.name + defValue;

      if (p.isOptional) {
        pName = `[${pName}]`;
      }

      str += linePrefix + `param {${pType}} ${pName}\n`;
    });
  }

  if (requirements) {
    requirements.forEach((requirement) => {
      str += linePrefix + `requires ${requirement}\n`;
    });
  }

  if (returns) {
    str += linePrefix + `returns {${returns}}\n`;
  }

  return startLine + str + endLine;
}

export function getCommentedModule(moduleText: string) {
  let output = "";

  output += `${getCommentedModuleOptionsSection(moduleText)}`;
  output += `${getCommentedModuleConstantsSection(moduleText)}`;
  output += `${getCommentedPrivateVariablesSection(moduleText)}`;
  output += `${getCommentedVBAEnumsSection(moduleText)}`;
  output += `${getCommentedProceduresSection(moduleText)}`;

  return output;
}

export function getCommentedModuleConstantsSection(moduleText: string) {
  let str = "";
  const modConsts = getModuleConstants(moduleText).sort((a, b) =>
    a.name.localeCompare(b.name)
  );

  modConsts.forEach((obj) => {
    str += `${obj.privacyType} Const ${obj.name} As ${obj.dataType} = ${obj.constValue}\n`;
  });

  if (str.length > 0) {
    str = `'MODULE CONSTANTS\n${str}\n`;
  }

  return str;
}

export function getCommentedModuleOptionsSection(moduleText: string) {
  let str = "";
  const modOptions = getModuleOptions(moduleText).sort();

  modOptions.forEach((obj) => {
    str += `${obj}\n`;
  });

  if (str.length > 0) {
    str = `'MODULE OPTIONS\n${str}\n`;
  }

  return str;
}

export function getCommentedPrivateVariablesSection(moduleText: string) {
  let str = "";
  const privateVars = getPrivateVariables(moduleText).sort((a, b) =>
    a.name.localeCompare(b.name)
  );

  privateVars.forEach((obj) => {
    str += `${obj.fullText}\n`;
  });

  if (str.length > 0) {
    str = `'PRIVATE VARIABLES\n${str}\n`;
  }

  return str;
}

export function getCommentedVBAEnumsSection(moduleText: string) {
  let str = "";
  const vbaEnums = getVBAEnums(moduleText).sort((a, b) =>
    a.name.localeCompare(b.name)
  );

  vbaEnums.forEach((obj) => {
    str += getCommentBlock({});
    str += `${obj.fullText}\n\n`;
  });

  if (str.length > 0) {
    str = `${str}\n`;
  }

  return str;
}

export function getCommentedProceduresSection(moduleText: string) {
  let str = "";
  const vbaProcedures = getVBAProcedures(moduleText)
    .sort(sortProceduresByName)
    .sort(sortProceduresByFunctionType);

  vbaProcedures.forEach((proc) => {
    str += getCommentBlock({
      params: proc.parameters,
      returns: proc.returnType,
      requirements: proc.requirements,
    });
    str += `${proc.fullText}\n\n`;
  });

  return str;
}

function sortProceduresByName(a: VBAProcedure, b: VBAProcedure) {
  // special case for 'Init' function
  if (a.name === "Init") return -1;

  return a.name.localeCompare(b.name);
}

function sortProceduresByFunctionType(a: VBAProcedure, b: VBAProcedure) {
  const functions = ["Function", "Sub"];
  const properties = ["Property Get", "Property Let", "Property Set"];

  // both in one or the other group
  if (
    functions.includes(a.functionType) == functions.includes(b.functionType)
  ) {
    return 0;
  }

  // a is property so should come after b
  return properties.includes(a.functionType) ? 1 : -1;
}
