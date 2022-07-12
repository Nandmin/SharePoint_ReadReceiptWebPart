/* eslint-disable @typescript-eslint/explicit-function-return-type */
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

// eslint-disable-next-line @typescript-eslint/no-unused-vars
const ReadReceiptWebpart: FunctionComponent<IReadReceiptWebpartProps> = (
  props
) => {
  const [showMessage, setShowMessage] = useState<boolean>(true);

  const {semanticColors}: IReadonlyTheme = props.themeVariant;

  useEffect(() => {
    if (props.storageList && props.storageList !== ""){
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      fetchData();
    }
  },  [props]);

  // eslint-disable-next-line @microsoft/spfx/no-async-await
  async function fetchData() {
    const items: unknown[] = await sp.web.lists
      .getById(props.storageList)
      .items.select("Author/ID", "Author/Title", "Author/Name", "Title")
      .expand("Author")
      .top(1)
      .filter(
        `Author/Title eq '${props.currentUserDisplayName}' and Title eq '${props.documentTitle}'`
      )
      .get();

    if (items.length === 0) {
      setShowMessage(true);
    }
  }

  const _onConfigure = () => {
    props.context.propertyPane.open();
  };

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  function _onChange(event: React.FormEvent<HTMLElement>, isChecked: boolean) {
    sp.web.lists.getById(props.storageList).items.add( {
      Title: props.documentTitle,
    });
    setShowMessage(false);
  }

  const mainStackTokens: IStackTokens = {
    childrenGap: 5,
    padding: 10,
  };

  return props.configured ? (
      <Stack style={{ backgroundColor: semanticColors.bodyBackground }}>
        { showMessage ? (
            <Stack
              style={{ color: semanticColors.bodyText }}
              tokens={mainStackTokens}>
                <Text>{props.acknowledgementMessage}</Text>
                <Text variant='large'>'{props.documentTitle}'</Text>
                <Checkbox
                  theme={props.themeVariant as ITheme}
                  label={props.acknowledgementLabel}
                  // eslint-disable-next-line react/jsx-no-bind
                  onChange={_onChange}
                  />
                    </Stack>
                    ) : (
                      <Stack style={{ color: semanticColors.bodyText}}>
                        <Text variant='large'>{props.documentTitle}</Text>
                        <Text>{props.readMessage}</Text>
                      </Stack>
                    )}
      </Stack>
          ):(
            <Placeholder
              iconName="Edit"
              iconText="Configure read receipt"
              description="Please configure..."
              buttonLabel="Configure"
              // eslint-disable-next-line react/jsx-no-bind
              onConfigure={_onConfigure}
              />
          );
};