"use client";

import { Switch } from "@fluentui/react-components";
import { WeatherMoonRegular, WeatherSunnyRegular } from "@fluentui/react-icons";
import { useAppTheme } from "./ThemeContext";

export function ThemeToggle() {
  const { setTheme, theme } = useAppTheme();
  const isDark = theme === "dark";

  return (
    <Switch
      checked={isDark}
      label={isDark ? "Dark" : "Light"}
      onChange={(_, data) => setTheme(data.checked ? "dark" : "light")}
      indicator={isDark ? <WeatherMoonRegular /> : <WeatherSunnyRegular />}
    />
  );
}
