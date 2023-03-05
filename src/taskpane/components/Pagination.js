import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton, Stack } from "@fluentui/react";

export default function Pagination({ pagination, next, prev }) {
  return (
    <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 40 }}>
      <DefaultButton onClick={prev} disabled={pagination.offset === 0}>
        Prev
      </DefaultButton>
      <DefaultButton onClick={next} disabled={pagination.offset + pagination.count >= pagination.total_count}>
        Next
      </DefaultButton>
    </Stack>
  );
}

Pagination.propTypes = {
  pagination: PropTypes.shape({
    total_count: PropTypes.number,
    count: PropTypes.number,
    offset: PropTypes.number,
  }),
  next: PropTypes.func,
  prev: PropTypes.func,
};
