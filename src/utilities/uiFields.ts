export type SelectOption = { value: string; text: string };

let fieldInfoTooltipCounter = 0;

function createFieldInfoTooltip(infoText: string, tooltipId: string): HTMLDivElement {
  const tooltip = document.createElement("div");
  tooltip.id = tooltipId;
  tooltip.className = "field__tooltip";
  tooltip.role = "tooltip";
  tooltip.hidden = true;
  tooltip.textContent = infoText;
  return tooltip;
}

function getFieldInfoHost(label: HTMLElement): HTMLElement | null {
  return (
    (label.closest(".date-range-filter__row") as HTMLElement | null) ??
    (label.closest(".field") as HTMLElement | null) ??
    (label.parentElement as HTMLElement | null)
  );
}

function positionFieldInfoTooltip(
  label: HTMLElement,
  tooltip: HTMLDivElement,
  tooltipHost: HTMLElement
): void {
  const topOffset = label.offsetTop + label.offsetHeight + 6;
  tooltip.style.top = `${topOffset}px`;
  tooltip.style.left = "0px";
  tooltip.style.right = "auto";
  tooltipHost.style.setProperty("--field-tooltip-top", `${topOffset}px`);
}

function createFieldInfoButton(infoText: string, tooltipId: string): HTMLButtonElement {
  const infoButton = document.createElement("button");
  infoButton.type = "button";
  infoButton.className = "field__info";
  infoButton.textContent = "i";
  infoButton.setAttribute("aria-label", infoText);
  infoButton.setAttribute("aria-describedby", tooltipId);
  infoButton.setAttribute("aria-expanded", "false");
  return infoButton;
}

export function attachFieldInfo(label: HTMLElement, infoText?: string): void {
  const normalizedInfoText = infoText?.trim();
  if (!normalizedInfoText || label.querySelector(".field__info")) {
    return;
  }

  const tooltipHost = getFieldInfoHost(label);
  if (!tooltipHost) {
    return;
  }

  const tooltipId = `field-info-tooltip-${++fieldInfoTooltipCounter}`;
  const infoButton = createFieldInfoButton(normalizedInfoText, tooltipId);
  const tooltip = createFieldInfoTooltip(normalizedInfoText, tooltipId);
  const showTooltip = () => {
    positionFieldInfoTooltip(label, tooltip, tooltipHost);
    infoButton.setAttribute("aria-expanded", "true");
    tooltip.hidden = false;
  };
  const hideTooltip = () => {
    infoButton.setAttribute("aria-expanded", "false");
    tooltip.hidden = true;
  };

  infoButton.addEventListener("mouseenter", showTooltip);
  infoButton.addEventListener("mouseleave", hideTooltip);
  infoButton.addEventListener("focus", showTooltip);
  infoButton.addEventListener("blur", hideTooltip);
  infoButton.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();

    if (tooltip.hidden) {
      showTooltip();
    } else {
      hideTooltip();
    }
  });

  label.appendChild(infoButton);
  tooltipHost.appendChild(tooltip);
}

export function createFieldWrapper(
  labelText: string,
  inputId: string,
  infoText?: string
): HTMLDivElement {
  const field = document.createElement("div");
  field.className = "field";

  const label = document.createElement("label");
  label.className = "field__label";
  label.htmlFor = inputId;
  label.textContent = labelText;
  field.appendChild(label);
  attachFieldInfo(label, infoText);

  return field;
}

export function createSelectField(args: {
  id: string;
  label: string;
  title?: string;
  infoText?: string;
  options: SelectOption[];
  value?: string;
}): HTMLDivElement {
  const infoText = args.infoText ?? args.title ?? args.label;
  const field = createFieldWrapper(args.label, args.id, infoText);

  const select = document.createElement("select");
  select.id = args.id;
  select.className = "select";
  if (infoText) select.title = infoText;

  for (const opt of args.options) {
    const o = document.createElement("option");
    o.value = opt.value;
    o.textContent = opt.text;
    select.appendChild(o);
  }

  if (args.value !== undefined) {
    select.value = args.value;
  }

  field.appendChild(select);
  return field;
}

export function createInputField(args: {
  id: string;
  label: string;
  title?: string;
  infoText?: string;
  type: "text" | "number";
  value?: string;
  placeholder?: string;
  min?: string;
  max?: string;
  step?: string;
}): HTMLDivElement {
  const infoText = args.infoText ?? args.title ?? args.label;
  const field = createFieldWrapper(args.label, args.id, infoText);

  const input = document.createElement("input");
  input.id = args.id;
  input.className = "input";
  input.type = args.type;
  if (infoText) input.title = infoText;
  if (args.placeholder) input.placeholder = args.placeholder;
  if (args.value !== undefined) input.value = args.value;
  if (args.min !== undefined) input.min = args.min;
  if (args.max !== undefined) input.max = args.max;
  if (args.step !== undefined) input.step = args.step;

  field.appendChild(input);
  return field;
}
