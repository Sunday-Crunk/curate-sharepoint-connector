declare interface IPreserveButtonCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'PreserveButtonCommandSetStrings' {
  const strings: IPreserveButtonCommandSetStrings;
  export = strings;
}
