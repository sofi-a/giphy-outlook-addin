import React, { useEffect, useState } from "react";
import PropTypes from "prop-types";
import { Image, ImageFit, Spinner, SpinnerSize, Stack } from "@fluentui/react";
import GIFs from "./GIFs";
import Pagination from "./Pagination";

const { GIPHY_API_KEY } = process.env;

export default function App({ isOfficeInitialized }) {
  const [gifs, setGifs] = useState([]);
  const [pagination, setPagination] = useState({
    total_count: 0,
    count: 25,
    offset: 0,
  });
  const [loading, setLoading] = useState(false);

  const fetchGifs = async ({ count, offset, q, endpoint = "trending" }) => {
    try {
      setLoading(true);
      const response = await fetch(
        `https://api.giphy.com/v1/gifs/${endpoint}?api_key=${GIPHY_API_KEY}&limit=${count}&offset=${offset}${
          endpoint === "search" ? `&q=${q}` : ""
        }`
      );
      const { data, pagination } = await response.json();
      setGifs(data);
      setPagination(pagination);
    } catch (error) {
      console.error(error);
    } finally {
      setLoading(false);
    }
  };

  const next = () => {
    setPagination((prev) => {
      fetchGifs({ count: prev.count, offset: prev.offset + prev.count });
      return { ...prev, offset: prev.offset + prev.count };
    });
    window.scrollTo(0, 0);
  };

  const prev = () => {
    setPagination((prev) => {
      fetchGifs({ count: prev.count, offset: prev.offset - prev.count });
      return { ...prev, offset: prev.offset - prev.count };
    });
    window.scrollTo(0, 0);
  };

  useEffect(() => {
    if (isOfficeInitialized) {
      fetchGifs({ count: pagination.count, offset: pagination.offset });
    }
  }, [isOfficeInitialized]);

  if (!isOfficeInitialized) {
    return (
      <Stack verticalAlign="center" style={{ height: "100vh" }}>
        <Spinner size={SpinnerSize.large} label="Initializing Add-in" />;
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <GIFs gifs={gifs} loading={loading} />
      <Pagination pagination={pagination} next={next} prev={prev} />
      <Image imageFit={ImageFit.contain} src={require("./../../../assets/attribution-mark.png")} />
    </Stack>
  );
}

App.propTypes = {
  isOfficeInitialized: PropTypes.bool,
};
