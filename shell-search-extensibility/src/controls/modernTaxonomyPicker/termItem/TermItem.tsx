import { IconButton } from '@fluentui/react/lib/components/Button/IconButton/IconButton';
import { classNamesFunction, styled } from '@fluentui/react/lib/Utilities';
import * as React from 'react';
import { ITermItemProps, ITermItemStyleProps, ITermItemStyles } from '../modernTermPicker/ModernTermPicker.types';
import { getStyles } from './TermItem.styles';
import styless from "./TermItem.module.scss";
import { IButtonStyles, IIconProps } from '@fluentui/react';

const getClassNames = classNamesFunction<ITermItemStyleProps, ITermItemStyles>();

/**
 * {@docCategory TagPicker}
 */
export const TermItemBase = (props: ITermItemProps) => {
  const {
    theme,
    styles,
    selected,
    disabled,
    enableTermFocusInDisabledPicker,
    children,
    className,
    index,
    onRemoveItem,
    removeButtonAriaLabel,
    termStoreInfo,
    languageTag,
  } = props;

  const classNames = getClassNames(styles, {
    theme: theme!,
    className,
    selected,
    disabled,
  });

  let labels = props.item.labels.filter((name) => name.languageTag === languageTag && name.isDefault);
  if (labels.length === 0) {
    labels = props.item.labels.filter((name) => name.languageTag === props.termStoreInfo.defaultLanguageTag && name.isDefault);
  }

  const addTermButtonStyles: IButtonStyles = { rootHovered: { backgroundColor: 'transparent' }, rootPressed: { backgroundColor: 'transparent' }, icon: { color: "#ffffff", cursor: "auto", margin: "0 0 0 12px" } };

  return (
    <div
      className={classNames.root}
      role={'listitem'}
      key={index}
      data-selection-index={index}
      data-is-focusable={(enableTermFocusInDisabledPicker || !disabled) && true}
      style={{ borderRadius: "9999px", height: "30px", lineHeight: "30px", backgroundColor: "#0078D4", color: "#ffffff" }}
    >
      <IconButton styles={addTermButtonStyles} iconProps={{ iconName: 'Tag' } as IIconProps} />
      <span className={classNames.text} aria-label={labels[0].name} title={labels[0].name}>
        {children}
      </span>
      <IconButton
        onClick={onRemoveItem}
        disabled={disabled}
        iconProps={{ iconName: 'Cancel', styles: { root: { fontSize: '12px' } } }}
        className={`${classNames.close} ${styless.term}`}
        ariaLabel={removeButtonAriaLabel}
        style={{ borderRadius: "9999px", color: "#ffffff" }}
      />
    </div>
  );
};

export const TermItem = styled<ITermItemProps, ITermItemStyleProps, ITermItemStyles>(TermItemBase, getStyles, undefined, {
  scope: 'TermItem',
});
