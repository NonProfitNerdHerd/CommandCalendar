import * as React from "react";
import { IMonacoEditorProps } from "./IMonacoEditorProps";

// CSP-safe editor: intentionally avoids loading Monaco/CDN scripts.
export const MonacoEditor: React.FunctionComponent<IMonacoEditorProps> = (
  props: React.PropsWithChildren<IMonacoEditorProps>
) => {
  const { value, onValueChange, readOnly } = props || ({} as IMonacoEditorProps);
  const [textValue, setTextValue] = React.useState<string>(value || "");

  React.useEffect(() => {
    setTextValue(value || "");
  }, [value]);

  const onTextareaChange = (e: React.ChangeEvent<HTMLTextAreaElement>): void => {
    const nextValue: string = e.target.value;
    setTextValue(nextValue);
    onValueChange(nextValue, []);
  };

  return (
    <>
      <div style={{ margin: "10px 0px", fontSize: "12px", color: "#605e5c" }}>
        Advanced editor running in CSP-safe textarea mode.
      </div>
      <textarea
        value={textValue}
        onChange={onTextareaChange}
        readOnly={readOnly}
        spellCheck={false}
        autoComplete="off"
        autoCorrect="off"
        autoCapitalize="none"
        rows={30}
        cols={115}
        style={{
          width: "100%",
          backgroundColor: "#111",
          color: "antiquewhite",
          padding: "12px",
          fontFamily: "Consolas, Monaco, monospace",
          fontSize: "12px"
        }}
      />
    </>
  );
};