import React, { useEffect, useState } from "react";
import PropTypes from "prop-types";
import { Image, ImageFit, SearchBox, Spinner, SpinnerSize, Stack } from "@fluentui/react";
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
  const [endpoint, setEndpoint] = useState("trending");
  const [searchTerm, setSearchTerm] = useState("");
  const [loading, setLoading] = useState(false);

  const fetchGifs = async ({ count, offset, q, endpoint }) => {
    try {
      setLoading(true);
      const response = await fetch(
        `https://api.giphy.com/v1/gifs/${endpoint}?api_key=${GIPHY_API_KEY}&limit=${count}&offset=${offset}${
          endpoint === "search" ? `&q=${q}` : ""
        }&lang=en`
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
      fetchGifs({
        count: prev.count,
        offset: prev.offset + prev.count,
        endpoint,
        ...(searchTerm && { q: searchTerm }),
      });
      return { ...prev, offset: prev.offset + prev.count };
    });
    window.scrollTo(0, 0);
  };

  const prev = () => {
    setPagination((prev) => {
      fetchGifs({
        count: prev.count,
        offset: prev.offset - prev.count,
        endpoint,
        ...(searchTerm && { q: searchTerm }),
      });
      return { ...prev, offset: prev.offset - prev.count };
    });
    window.scrollTo(0, 0);
  };

  const onSearch = (searchTerm) => {
    if (searchTerm) {
      setEndpoint("search");
      setPagination({ total_count: 0, count: 25, offset: 0 });
      fetchGifs({ count: pagination.count, offset: 0, q: searchTerm, endpoint: "search" });
    }
  };

  const onClear = () => {
    setSearchTerm("");
    setEndpoint("trending");
    setPagination({ total_count: 0, count: 25, offset: 0 });
    fetchGifs({ count: pagination.count, offset: 0, endpoint: "trending" });
  };

  const insertGif = (id) => {
    const src = gifs.find((gif) => gif.id === id).images.original.url;
    Office.context.mailbox.item.body.setSelectedDataAsync(
      `<div><img src="${src}" /></div>`,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(asyncResult.error.message);
        }
      }
    );
  };

  useEffect(() => {
    if (isOfficeInitialized) {
      fetchGifs({ count: pagination.count, offset: pagination.offset, endpoint });
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
      <SearchBox
        placeholder="Search"
        underlined
        value={searchTerm}
        onChange={(e) => {
          setSearchTerm(e.target.value);
          onSearch(e.target.value);
        }}
        onSearch={onSearch}
        onClear={onClear}
      />
      <GIFs gifs={gifs} loading={loading} onClick={insertGif} />
      <Pagination pagination={pagination} next={next} prev={prev} />
      <Image imageFit={ImageFit.contain} src={require("./../../../assets/attribution-mark.png")} />
    </Stack>
  );
}

App.propTypes = {
  isOfficeInitialized: PropTypes.bool,
};
