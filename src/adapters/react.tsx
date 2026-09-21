import { createElement, useEffect, useRef, type CSSProperties } from "react";
import { createEditorSession, type AdapterProps } from "./session";
export type { EditorHandle, AdapterProps } from "./session";
export interface OnlineExcelProps extends AdapterProps {
  className?: string;
  style?: CSSProperties;
}
/** A thin view adapter. Set a height on this component or its parent. */
export function OnlineExcel(props: OnlineExcelProps) {
  const container = useRef<HTMLDivElement>(null);
  const latest = useRef(props);
  latest.current = props;
  const session = useRef<ReturnType<typeof createEditorSession> | undefined>(
    undefined,
  );
  useEffect(() => {
    const current = createEditorSession(
      container.current!,
      latest.current,
      () => latest.current,
    );
    session.current = current;
    return () => {
      current.dispose();
      if (session.current === current) session.current = undefined;
    };
  }, [props.workbook]);
  useEffect(() => {
    session.current?.update();
  }, [props.options]);
  return createElement("div", {
    ref: container,
    className: props.className,
    style: { height: "100%", ...props.style },
  });
}
