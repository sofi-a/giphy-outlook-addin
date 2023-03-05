import * as React from "react";
import PropTypes from "prop-types";
import { Image, ImageFit, Spinner, SpinnerSize, Stack } from "@fluentui/react";
import GIFs from "./GIFs";
import gifs from "./gifs.json";

/* global require */

export default function App({ isOfficeInitialized }) {
  if (!isOfficeInitialized) {
    return (
      <Stack verticalAlign="center" style={{ height: "100vh" }}>
        <Spinner size={SpinnerSize.large} label="Initializing Add-in" />;
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <GIFs gifs={gifs} />
      <Image imageFit={ImageFit.contain} src={require("./../../../assets/attribution-mark.png")} />
    </Stack>
  );
}

App.propTypes = {
  isOfficeInitialized: PropTypes.bool,
};
