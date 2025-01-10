import {
  ModuleConstant,
  ModuleConstantRegexGroup,
  PassByType,
  Privacy,
  PrivateVariable,
  PrivateVariableRegexGroup,
  VBAEnum,
  VBAEnumRegexGroup,
  VBAProcedureWithNoReturnRegexGroup,
  VBAProcedureWithReturnRegexGroup,
  VBAParam,
  VBAParamRegexGroup,
  VBAProcedure,
  VBAEnumValuesRegexGroup,
} from "./types";

export function getModuleConstants(moduleText: string) {
  //match module level consts only
  //only works as procedure consts are indented
  const constRegex =
    /^(?:(Public|Private)\s)?Const\s(\w+)\sAs\s(\w+)\s=\s((?:.)+)$/gm;
  let regexMatch: RegExpExecArray | null = null;
  let constantList: ModuleConstant[] = [];

  do {
    regexMatch = constRegex.exec(moduleText);
    if (regexMatch) {
      let privacyType: Privacy;

      privacyType =
        regexMatch[ModuleConstantRegexGroup.privacyType] == Privacy.private
          ? Privacy.private
          : Privacy.public;

      constantList.push({
        privacyType,
        name: regexMatch[ModuleConstantRegexGroup.name],
        dataType: regexMatch[ModuleConstantRegexGroup.dataType],
        constValue: regexMatch[ModuleConstantRegexGroup.constValue],
        fullText: regexMatch[ModuleConstantRegexGroup.fullText],
      });
    }
  } while (regexMatch);

  return constantList;
}

export function getModuleOptions(moduleText: string) {
  //Option Explicit or Option Private Module
  const optionsRegex = /^Option\s(?:Explicit|Private\sModule)$/gm;
  let regexMatch: RegExpExecArray | null = null;
  let optionsList: string[] = [];

  do {
    regexMatch = optionsRegex.exec(moduleText);
    if (regexMatch) {
      optionsList.push(regexMatch[0]);
    }
  } while (regexMatch);

  return optionsList;
}

export function getPrivateVariables(moduleText: string) {
  const privateVarRegex = /^Private\s(\w+)\sAs\s(\w+)$/gm;
  let regexMatch: RegExpExecArray | null = null;
  let privateVarList: PrivateVariable[] = [];

  do {
    regexMatch = privateVarRegex.exec(moduleText);
    if (regexMatch) {
      privateVarList.push({
        name: regexMatch[PrivateVariableRegexGroup.name],
        dataType: regexMatch[PrivateVariableRegexGroup.dataType],
        fullText: regexMatch[PrivateVariableRegexGroup.fullText],
      });
    }
  } while (regexMatch);

  return privateVarList;
}

export function getVBAEnums(moduleText: string) {
  const vbaEnumRegex =
    /^(?:(Public|Private)\s)?Enum\s(\w+)((?:.|\n)+?)End Enum$/gm;
  let regexMatch: RegExpExecArray | null = null;
  let vbaEnumList: VBAEnum[] = [];

  do {
    regexMatch = vbaEnumRegex.exec(moduleText);
    if (regexMatch) {
      let privacyType: Privacy;

      privacyType =
        regexMatch[VBAEnumRegexGroup.privacyType] == Privacy.private
          ? Privacy.private
          : Privacy.public;

      let valueMap = getVBAEnumsValueMap(regexMatch[VBAEnumRegexGroup.body]);

      vbaEnumList.push({
        privacyType: privacyType,
        name: regexMatch[VBAEnumRegexGroup.name],
        fullText: regexMatch[VBAEnumRegexGroup.fullText],
        body: regexMatch[VBAEnumRegexGroup.body],
        valueMap: valueMap,
      });
    }
  } while (regexMatch);

  return vbaEnumList;
}

function getVBAEnumsValueMap(enumBodyText: string) {
  const enumValuesRegex = /(\w+)(?:\s=\s(\d+))?/gm;
  let regexMatch: RegExpExecArray | null = null;
  let enumValMap: [string, number][] = [];

  do {
    regexMatch = enumValuesRegex.exec(enumBodyText);
    if (regexMatch) {
      let name = regexMatch[VBAEnumValuesRegexGroup.name];
      let valueMatch = regexMatch[VBAEnumValuesRegexGroup.value];

      // if no value grab last value and add 1
      let value = 0;

      if (valueMatch) {
        value = parseInt(valueMatch);
      } else {
        if (enumValMap.length > 0) {
          value = enumValMap[enumValMap.length - 1][1] + 1;
        }
      }
      enumValMap.push([name, value]);
    }
  } while (regexMatch);

  return enumValMap;
}

export function getVBAProcedureRequirementsByWorkbook(
  wbNames: string[],
  procedureBody: string
) {
  let wbStr = "";

  wbNames.forEach((wbName) => {
    wbStr += wbName + "|";
  });

  //remove trailing '|'
  wbStr = wbStr.substring(0, wbStr.length - 1);

  let reg = new RegExp(`(${wbStr})(?:\\.(\\w+))(?:\\.(\\w+))?`, "g");
  let regexMatch: RegExpExecArray | null = null;
  let requirements = new Set<string>();

  do {
    regexMatch = reg.exec(procedureBody);

    if (regexMatch) {
      requirements.add(regexMatch[0]);
    }
  } while (regexMatch);

  return requirements;
}

export function getVBAProcedures(moduleText: string) {
  const vbaProcedureWithReturnRegex =
    /^(?:(Public|Private)\s)?(Function|Property Get)\s(\w+)\(((?:.|\n)*?)\)\sAs\s(\w+)\n((?:.|\n)+?)End\s(?:Function|Property)$/gm;
  const vbaProcedureWithNoReturnRegex =
    /^(?:(Public|Private)\s)?(Sub|Property Let|Property Set)\s(\w+)\(((?:.|\n)*?)\)\n((?:.|\n)+?)End\s(?:Sub|Property)$/gm;
  let regexMatch: RegExpExecArray | null = null;
  let vbaProcedures: VBAProcedure[] = [];

  //get functions + property get
  do {
    regexMatch = vbaProcedureWithReturnRegex.exec(moduleText);
    if (regexMatch) {
      let privacyType: Privacy;

      privacyType =
        regexMatch[VBAProcedureWithReturnRegexGroup.privacyType] ==
        Privacy.private
          ? Privacy.private
          : Privacy.public;

      let params = GetParameters(
        regexMatch[VBAProcedureWithReturnRegexGroup.parameters]
      );

      let requirements = getVBAProcedureRequirementsByWorkbook(
        ["MyAddIn"],
        regexMatch[VBAProcedureWithReturnRegexGroup.body]
      );

      vbaProcedures.push({
        privacyType: privacyType,
        name: regexMatch[VBAProcedureWithReturnRegexGroup.name],
        fullText: regexMatch[VBAProcedureWithReturnRegexGroup.fullText],
        body: regexMatch[VBAProcedureWithReturnRegexGroup.body],
        functionType: regexMatch[VBAProcedureWithReturnRegexGroup.functionType],
        returnType: regexMatch[VBAProcedureWithReturnRegexGroup.returnType],
        parameters: params,
        requirements: Array.from(requirements),
      });
    }
  } while (regexMatch);

  //get subs + property [let|set]
  do {
    regexMatch = vbaProcedureWithNoReturnRegex.exec(moduleText);
    if (regexMatch) {
      let privacyType: Privacy;

      privacyType =
        regexMatch[VBAProcedureWithNoReturnRegexGroup.privacyType] ==
        Privacy.private
          ? Privacy.private
          : Privacy.public;

      let params = GetParameters(
        regexMatch[VBAProcedureWithNoReturnRegexGroup.parameters]
      );

      vbaProcedures.push({
        privacyType: privacyType,
        name: regexMatch[VBAProcedureWithNoReturnRegexGroup.name],
        fullText: regexMatch[VBAProcedureWithNoReturnRegexGroup.fullText],
        body: regexMatch[VBAProcedureWithNoReturnRegexGroup.body],
        functionType:
          regexMatch[VBAProcedureWithNoReturnRegexGroup.functionType],
        parameters: params,
      });
    }
  } while (regexMatch);

  return vbaProcedures;
}

export function GetParameters(paramsText: string) {
  const vbaParametersRegex =
    /(?:(Optional)\s)?(?:(ByVal|ByRef)\s)?(?:(ParamArray)\s)?(\w+)(?:\(\))?(?:\sAs\s(\w+))?(?:\s=\s([^\),]+))?/gm;
  let regexMatch: RegExpExecArray | null = null;
  let vbaParameters: VBAParam[] = [];

  // FIXME case for no parameters (can't figure out regex to match 0 and 1+ params at same time)
  if (paramsText.charAt(0) === ")") return vbaParameters;

  do {
    regexMatch = vbaParametersRegex.exec(paramsText);
    if (regexMatch) {
      // FIXME special case for multi-line parameters (would return each '_' as own parameter)
      if (!(regexMatch[VBAParamRegexGroup.name] === "_")) {
        vbaParameters.push({
          dataType: regexMatch[VBAParamRegexGroup.type] ?? "Variant",
          isParamArray: regexMatch[VBAParamRegexGroup.isParamArray]
            ? true
            : false,
          isOptional: regexMatch[VBAParamRegexGroup.isOptional] ? true : false,
          name: regexMatch[VBAParamRegexGroup.name],
          passByType:
            regexMatch[VBAParamRegexGroup.passByType] == PassByType.byVal
              ? PassByType.byVal
              : PassByType.byRef,
          defaultValue: regexMatch[VBAParamRegexGroup.defaultValue],
        });
      }
    }
  } while (regexMatch);

  return vbaParameters;
}
