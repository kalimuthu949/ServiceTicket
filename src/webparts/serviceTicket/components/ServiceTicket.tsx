import * as React from 'react';
import { IServiceTicketProps } from './IServiceTicketProps';
import { escape } from '@microsoft/sp-lodash-subset';
import "../../../ExternalRef/css/style.css";
import { sp } from "@pnp/sp";
import MainServiceTicket from './MainServiceTicket';

export default class ServiceTicket extends React.Component<IServiceTicketProps, {}> {
  constructor(prop: IServiceTicketProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
  }
  public render(): React.ReactElement<IServiceTicketProps> {
    return <MainServiceTicket context={this.props.context} sp={sp} />;
  }
}
