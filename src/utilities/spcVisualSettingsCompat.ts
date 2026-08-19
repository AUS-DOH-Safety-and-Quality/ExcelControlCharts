import extractConditionalFormatting from "../PowerBI-SPC/src/Functions/extractConditionalFormatting";
import {
  defaultSettings as spcDefaultSettings,
  type settingsValueType as spcSettingsValueType,
} from "../PowerBI-SPC/src/settings";

type SpcInputSettingsInstance = {
  settings?: spcSettingsValueType[];
  derivedSettings?: Array<{ update: (spcSettings: spcSettingsValueType["spc"]) => void }>;
  update?: (inputView: unknown, groupIdxs: number[][]) => void;
  __taskpanePatched?: boolean;
};

function backfillMissingSpcSettings(
  inputSettings: SpcInputSettingsInstance,
  inputView: any,
  groupIdxs: number[][]
) {
  if (!Array.isArray(inputSettings.settings)) {
    return;
  }

  const settingGroups = Object.keys(spcDefaultSettings) as Array<keyof spcSettingsValueType>;

  settingGroups.forEach((settingGroup) => {
    const defaultGroupSettings = spcDefaultSettings[settingGroup] as Record<string, unknown>;
    const defaultGroupSettingNames = Object.keys(defaultGroupSettings);

    groupIdxs.forEach((groupRows, groupIndex) => {
      const currentSettings = inputSettings.settings?.[groupIndex];
      if (!currentSettings) {
        return;
      }

      const currentGroupSettings = ((currentSettings as any)[settingGroup] ?? {}) as Record<
        string,
        unknown
      >;
      const missingSettingNames = defaultGroupSettingNames.filter(
        (settingName) => !Object.prototype.hasOwnProperty.call(currentGroupSettings, settingName)
      );

      if (missingSettingNames.length === 0) {
        return;
      }

      const extractedGroupSettings = extractConditionalFormatting<any>(
        inputView?.categorical,
        settingGroup as string,
        spcDefaultSettings,
        groupRows
      )?.values?.[0] as Record<string, unknown> | undefined;

      const nextGroupSettings: Record<string, unknown> = {
        ...defaultGroupSettings,
        ...currentGroupSettings,
      };

      missingSettingNames.forEach((settingName) => {
        nextGroupSettings[settingName] =
          extractedGroupSettings?.[settingName] ?? defaultGroupSettings[settingName];
      });

      (currentSettings as any)[settingGroup] = nextGroupSettings;
    });
  });

  inputSettings.settings.forEach((settingsItem, groupIndex) => {
    inputSettings.derivedSettings?.[groupIndex]?.update(settingsItem.spc);
  });
}

function patchSpcVisualSettingsForTaskpane(spcVisual: any) {
  const inputSettings = spcVisual?.viewModel?.inputSettings as SpcInputSettingsInstance | undefined;
  if (!inputSettings || typeof inputSettings.update !== "function" || inputSettings.__taskpanePatched) {
    return;
  }

  inputSettings.__taskpanePatched = true;
  const originalUpdate = inputSettings.update.bind(inputSettings);

  inputSettings.update = ((inputView: unknown, groupIdxs: number[][]) => {
    originalUpdate(inputView, groupIdxs);
    backfillMissingSpcSettings(inputSettings, inputView, groupIdxs);
  }) as SpcInputSettingsInstance["update"];
}

export { backfillMissingSpcSettings, patchSpcVisualSettingsForTaskpane };