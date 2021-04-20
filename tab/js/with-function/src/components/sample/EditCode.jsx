import React from "react";

var functionName = process.env.REACT_APP_FUNC_NAME || "myFunc";
export function EditCode(props) {
  const { showFunction } = {
    showFunction: true,
    ...props,
  };
  return (
    <div>
      <h2>Change this code</h2>
      <p>
        The front end is a <code>create-react-app</code>. The entry point is <code>tabs/src/index.js</code>. Just save any file and this page will reload automatically.
      </p>
      {showFunction && (
        <p>
          This app contains an Azure Functions backend. Find the code in{" "}
          <code>api/{functionName}/index.js</code>
        </p>
      )}
    </div>
  );
}
