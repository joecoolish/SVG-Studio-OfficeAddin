import * as React from "react";
import { Button, makeStyles, Text, Tooltip, tokens } from "@fluentui/react-components";
import { WeatherMoonRegular, WeatherSunnyRegular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  root: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "12px",
    flexShrink: 0,
  },
  brand: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    minWidth: 0,
  },
  logo: {
    width: "36px",
    height: "36px",
    objectFit: "contain",
    flexShrink: 0,
  },
  text: {
    display: "flex",
    flexDirection: "column",
    lineHeight: 1.2,
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
  },
  subtitle: {
    color: tokens.colorNeutralForeground3,
  },
});

interface HeaderProps {
  darkMode: boolean;
  onToggleTheme: () => void;
}

const Header: React.FC<HeaderProps> = ({ darkMode, onToggleTheme }) => {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <div className={styles.brand}>
        <img className={styles.logo} src="assets/icon-64.png?v=5" alt="" />
        <div className={styles.text}>
          <Text size={500} className={styles.title}>
            SVG Studio
          </Text>
          <Text size={200} className={styles.subtitle}>
            Edit &amp; insert
          </Text>
        </div>
      </div>
      <Tooltip content={`Switch to ${darkMode ? "light" : "dark"} mode`} relationship="label">
        <Button
          appearance="subtle"
          icon={darkMode ? <WeatherSunnyRegular /> : <WeatherMoonRegular />}
          aria-label={`Switch to ${darkMode ? "light" : "dark"} mode`}
          onClick={onToggleTheme}
        />
      </Tooltip>
    </div>
  );
};

export default Header;
