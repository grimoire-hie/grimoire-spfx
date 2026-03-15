/**
 * PasswordField
 * Custom property pane field for secure API key input with reveal toggle
 */

import * as React from 'react';
import { TextField, IconButton } from '@fluentui/react';

export interface IPropertyPanePasswordFieldProps {
  label: string;
  value: string;
  description?: string;
  onChange: (value: string) => void;
}

export const PropertyPanePasswordField: React.FC<IPropertyPanePasswordFieldProps> = ({
  label,
  value,
  description,
  onChange
}) => {
  const [revealed, setRevealed] = React.useState<boolean>(false);

  const handleChange = (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    onChange(newValue || '');
  };

  const toggleReveal = (): void => {
    setRevealed(prev => !prev);
  };

  return (
    <TextField
      label={label}
      value={value}
      type={revealed ? 'text' : 'password'}
      onChange={handleChange}
      description={description}
      autoComplete="off"
      onRenderSuffix={() => (
        <IconButton
          iconProps={{ iconName: revealed ? 'Hide' : 'RedEye' }}
          ariaLabel={revealed ? 'Hide value' : 'Show value'}
          onClick={toggleReveal}
          styles={{ root: { marginRight: -8 } }}
        />
      )}
    />
  );
};
