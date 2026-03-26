export enum GroupingStrategy {
  Date = "Date",
  Text = "Text",
  Choice = "Choice",
}

export const GroupingStrategyLabels: Record<GroupingStrategy, string> = {
  [GroupingStrategy.Date]: "By Date",
  [GroupingStrategy.Text]: "By Initial Letter",
  [GroupingStrategy.Choice]: "By Choice Value",
};

export const MaxLevelsForStrategy: Record<GroupingStrategy, number> = {
  [GroupingStrategy.Date]: 3,
  [GroupingStrategy.Text]: 3,
  [GroupingStrategy.Choice]: 1,
};
