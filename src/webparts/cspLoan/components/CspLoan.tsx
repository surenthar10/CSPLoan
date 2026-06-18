import * as React from "react";
// import styles from "./CspLoan.module.scss";
import type { ICspLoanProps } from "./ICspLoanProps";
import "primereact/resources/themes/bootstrap4-light-blue/theme.css";
import MainComponent from "./MainComponent";

export default class CspLoan extends React.Component<ICspLoanProps, {}> {
  public render(): React.ReactElement<ICspLoanProps> {
    return <MainComponent context={this.props.context} />;
  }
}
