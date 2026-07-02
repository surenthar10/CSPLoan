import * as React from "react";
// import styles from "./CspLoan.module.scss";
import type { ICspLoanProps } from "./ICspLoanProps";
import "primereact/resources/themes/bootstrap4-light-blue/theme.css";
import MainComponent from "./MainComponent";
import { Provider } from "react-redux";
import store from "../assets/Redux/Store/store";

export default class CspLoan extends React.Component<ICspLoanProps, {}> {
  public render(): React.ReactElement<ICspLoanProps> {
    return (
      <Provider store={store}>
        <MainComponent context={this.props.context} />
      </Provider>
    );
  }
}
