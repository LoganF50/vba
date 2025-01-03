export enum PassByType {
  byRef = "ByRef",
  byVal = "ByVal",
}

export enum Privacy {
  private = "Private",
  public = "Public",
}

export type ModuleConstant = {
  privacyType: Privacy;
  name: string;
  dataType: string;
  constValue: string | number | boolean;
  fullText: string;
};

export enum ModuleConstantRegexGroup {
  fullText = 0,
  privacyType = 1,
  name = 2,
  dataType = 3,
  constValue = 4,
}

export type PrivateVariable = {
  name: string;
  dataType: string;
  fullText: string;
};

export enum PrivateVariableRegexGroup {
  fullText = 0,
  name = 1,
  dataType = 2,
}

export type VBAClassModule = {
  name: string;
  procedures?: VBAProcedure[];
  dependents?: string[];
  requirements?: string[];
};

export type VBAEnum = {
  privacyType: Privacy;
  name: string;
  fullText: string;
  body: string;
  description?: string;
  dependents?: string[];
  valueMap?: [string, number][];
};

export enum VBAEnumRegexGroup {
  fullText = 0,
  privacyType = 1,
  name = 2,
  body = 3,
}

export enum VBAEnumValuesRegexGroup {
  fullText = 0,
  name = 1,
  value = 2,
}

export type VBAForm = {
  name: string;
  procedures?: VBAProcedure[];
  dependents?: string[];
};

export type VBAModule = {
  name: string;
  moduleConstants?: ModuleConstant[];
  privateVariables?: PrivateVariable[];
  enums?: VBAEnum[];
  procedures?: VBAProcedure[];
};

export type VBAParam = {
  dataType: string;
  isParamArray: boolean;
  isOptional: boolean;
  name: string;
  defaultValue?: string | number | boolean;
  passByType: PassByType;
  description?: string;
};

export enum VBAParamRegexGroup {
  fullText = 0,
  isOptional = 1,
  passByType = 2,
  isParamArray = 3,
  name = 4,
  type = 5,
  defaultValue = 6,
}

export type VBAProcedure = {
  privacyType: Privacy;
  functionType: string;
  name: string;
  parameters: VBAParam[];
  returnType?: string;
  fullText: string;
  body: string;
  description?: string;
  dependents?: string[];
  needs?: string[];
  requirements?: string[];
};

export enum VBAProcedureWithReturnRegexGroup {
  fullText = 0,
  privacyType = 1,
  functionType = 2,
  name = 3,
  parameters = 4,
  returnType = 5,
  body = 6,
}

export enum VBAProcedureWithNoReturnRegexGroup {
  fullText = 0,
  privacyType = 1,
  functionType = 2,
  name = 3,
  parameters = 4,
  body = 5,
}

export type VBAWorkbook = {
  name: string;
  classModules?: VBAClassModule[];
  forms?: VBAForm[];
  vbaModules?: VBAModule[];
};
