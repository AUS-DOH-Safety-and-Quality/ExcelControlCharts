export type SelectOption = { value: string; text: string };

export function createFieldWrapper(labelText: string, inputId: string): HTMLDivElement {
  const field = document.createElement("div");
  field.className = "field";

  const label = document.createElement("label");
  label.htmlFor = inputId;
  label.textContent = labelText;
  field.appendChild(label);

  return field;
}

export function createSelectField(args: {
  id: string;
  label: string;
  title?: string;
  options: SelectOption[];
  value?: string;
}): HTMLDivElement {
  const field = createFieldWrapper(args.label, args.id);

  const select = document.createElement("select");
  select.id = args.id;
  select.className = "select";
  if (args.title) select.title = args.title;

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
  type: "text" | "number";
  value?: string;
  placeholder?: string;
  min?: string;
  max?: string;
  step?: string;
}): HTMLDivElement {
  const field = createFieldWrapper(args.label, args.id);

  const input = document.createElement("input");
  input.id = args.id;
  input.className = "input";
  input.type = args.type;
  if (args.title) input.title = args.title;
  if (args.placeholder) input.placeholder = args.placeholder;
  if (args.value !== undefined) input.value = args.value;
  if (args.min !== undefined) input.min = args.min;
  if (args.max !== undefined) input.max = args.max;
  if (args.step !== undefined) input.step = args.step;

  field.appendChild(input);
  return field;
}
