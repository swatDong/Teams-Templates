import React, { ReactElement } from "react";
import classnames from "classnames";
import { AcceptIcon } from "@fluentui/react-northstar";
import "./Progress.css";

export function Progress(props: { children?: ReactElement[], selectedIndex: number }) {
  return (
    <div className="progress-indicator">
      <div className="line"></div>
      {React.Children.map(props.children, (child, i) => (
        <div
          className={classnames("progress-item", {
            selected: props.selectedIndex === i,
          })}
          key={i}
        >
          <div className={classnames("check")}>
            {props.selectedIndex === i && <AcceptIcon size="smaller" />}
          </div>
          <div className="content">{child}</div>
        </div>
      ))}
    </div>
  );
}
