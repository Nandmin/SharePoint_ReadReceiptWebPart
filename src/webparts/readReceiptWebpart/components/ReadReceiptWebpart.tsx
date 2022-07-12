import * as React from 'react';
import styles from './ReadReceiptWebpart.module.scss';
import { IReadReceiptWebpartProps } from './IReadReceiptWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { FunctionComponent, useEffect, useState } from 'react';
import { sp } from "@pnp/sp"
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Checkbox, Text, IStackTokens, ITheme, Stack} from "office-ui-fabric-react";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';

const ReadReceiptWebpart: FunctionComponent<IReadReceiptWebpartProps> = (
  props
) => {
  
}
