import { createAvatarSettingsDropdownStyles } from './AvatarSettingsPanel';

interface ITestStyle {
  backgroundColor?: string;
  border?: string;
  color?: string;
  selectors?: Record<string, { color?: string; backgroundColor?: string }>;
}

describe('AvatarSettingsPanel dropdown visibility', () => {
  it('maps dropdown visuals to provided SharePoint theme colors', () => {
    const styles = createAvatarSettingsDropdownStyles({
      bodyBackground: '#faf9f8',
      bodyText: '#323130',
      bodySubtext: '#605e5c',
      cardBackground: '#ffffff',
      cardBorder: '#edebe9',
      isDark: false
    });

    const titleStyle = styles.title as ITestStyle | undefined;
    const calloutStyle = styles.callout as ITestStyle | undefined;
    const itemStyle = styles.dropdownItem as ITestStyle | undefined;
    const selectedStyle = styles.dropdownItemSelected as ITestStyle | undefined;

    expect(titleStyle?.backgroundColor).toBe('#faf9f8');
    expect(titleStyle?.color).toBe('#323130');
    expect(titleStyle?.border).toBe('1px solid #edebe9');
    expect(titleStyle?.selectors?.['.ms-Dropdown-titleText']?.color).toContain('#323130');

    expect(calloutStyle?.backgroundColor).toBe('#ffffff');
    expect(calloutStyle?.border).toBe('1px solid #edebe9');

    expect(itemStyle?.color).toBe('#323130');
    expect(selectedStyle?.color).toBe('#323130');
  });
});
