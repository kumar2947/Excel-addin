import * as React from "react";
import PropTypes from "prop-types";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = (props) => {
  const styles = useStyles();


  return (
    <div className={styles.root}>
     
      <TextInsertion />
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
