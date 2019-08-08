declare interface ICustomCommandCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'CustomCommandCommandSetStrings' {
  const strings: ICustomCommandCommandSetStrings;
  export = strings;
}
