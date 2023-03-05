import * as React from "react";
import PropTypes from "prop-types";
import { FontIcon, Image, ImageFit, Link, Shimmer, ShimmerElementType, Stack, Text } from "@fluentui/react";

export default function GIFs({ gifs = [], loading }) {
  return (
    <div className="ms-welcome__features ms-u-fadeIn500">
      <div className="ms-Grid" dir="ltr">
        <div className="ms-Grid-row">
          {loading
            ? Array.from(Array(4).keys()).map((i) => (
                <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg4 ms-xl3" style={{ marginBottom: "2rem" }}>
                  <Stack tokens={{ childrenGap: 5 }}>
                    <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 200 }]} />
                    <Shimmer
                      shimmerElements={[
                        { type: ShimmerElementType.circle, height: 35 },
                        { type: ShimmerElementType.gap },
                      ]}
                    />
                  </Stack>
                </div>
              ))
            : gifs.map((gif) => (
                <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg4 ms-xl3" style={{ marginBottom: "2rem" }}>
                  <Stack key={gif.id} tokens={{ childrenGap: 5 }}>
                    <Image
                      imageFit={ImageFit.contain}
                      alt={gif.title}
                      title={gif.title}
                      src={gif.images.fixed_height.url}
                      height={gif.images.fixed_height.height}
                      style={{ cursor: "pointer" }}
                    />
                    {gif.user && (
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <Image src={gif.user.avatar_url} width={35} style={{ borderRadius: "50%" }} />
                        <Stack>
                          <Text variant="large" style={{ fontWeight: "bold" }}>
                            {gif.user.display_name}
                          </Text>
                          <Text variant="medium">
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 5 }}>
                              <Link href={gif.user.profile_url} target="_blank" style={{ textDecoration: "none" }}>
                                @{gif.username}
                              </Link>
                              {gif.user.is_verified && (
                                <FontIcon
                                  aria-label="Verified"
                                  iconName="VerifiedBrandSolid"
                                  style={{ color: "#15cdff" }}
                                />
                              )}
                            </Stack>
                          </Text>
                        </Stack>
                      </Stack>
                    )}
                  </Stack>
                </div>
              ))}
        </div>
      </div>
    </div>
  );
}

GIFs.propTypes = {
  gifs: PropTypes.arrayOf(
    PropTypes.shape({
      id: PropTypes.string,
      title: PropTypes.string,
      images: PropTypes.shape({
        fixed_height: PropTypes.shape({
          url: PropTypes.string,
          height: PropTypes.number,
        }),
      }),
      user: PropTypes.shape({
        avatar_url: PropTypes.string,
        display_name: PropTypes.string,
        profile_url: PropTypes.string,
        is_verified: PropTypes.bool,
      }),
    })
  ),
  loading: PropTypes.bool,
};
