import type { MessageSnapshot } from "./types";

export type SurfaceMode = "compose" | "read" | "unknown";

function getMailboxItem(): Office.MessageCompose | Office.MessageRead | undefined {
  return Office.context?.mailbox?.item as Office.MessageCompose | Office.MessageRead | undefined;
}

function isComposeItem(item: Office.MessageCompose | Office.MessageRead | undefined): item is Office.MessageCompose {
  return Boolean(item && "body" in item && typeof (item as Office.MessageCompose).body?.setSelectedDataAsync === "function");
}

function getAsync<T>(executor: (callback: (result: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) => {
    executor((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }

      reject(new Error(result.error?.message ?? "Office.js operation failed."));
    });
  });
}

function setAsync<T>(executor: (callback: (result: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return getAsync(executor);
}

export function getSurfaceMode(): SurfaceMode {
  const item = getMailboxItem();

  if (!item) {
    return "unknown";
  }

  return isComposeItem(item) ? "compose" : "read";
}

export async function getCurrentMessageSnapshot(): Promise<MessageSnapshot> {
  const item = getMailboxItem();

  if (!item) {
    return {
      bodyText: "",
      subject: "",
      to: [],
      mode: "unknown",
    };
  }

  if (isComposeItem(item)) {
    const [body, subject, recipients] = await Promise.all([
      getAsync<string>((callback) => item.body.getAsync(Office.CoercionType.Text, callback)),
      getAsync<string>((callback) => item.subject.getAsync(callback)),
      getAsync<Office.EmailAddressDetails[]>((callback) => item.to.getAsync(callback)),
    ]);

    return {
      bodyText: body ?? "",
      subject: subject ?? "",
      to: recipients.map((recipient) => recipient.emailAddress).filter(Boolean),
      conversationId: item.conversationId,
      mode: "compose",
    };
  }

  const readItem = item as Office.MessageRead;
  const body = await getAsync<string>((callback) => readItem.body.getAsync(Office.CoercionType.Text, callback));

  return {
    bodyText: body ?? "",
    subject: readItem.subject ?? "",
    to: [],
    from: readItem.from?.emailAddress,
    conversationId: readItem.conversationId,
    mode: "read",
  };
}

export async function insertTextIntoCompose(text: string): Promise<void> {
  const item = getMailboxItem();

  if (!isComposeItem(item)) {
    throw new Error("Open a draft or reply compose window before inserting text.");
  }

  await setAsync<void>((callback) =>
    item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, callback),
  );
}

export async function prependBodyText(text: string): Promise<void> {
  const item = getMailboxItem();

  if (!isComposeItem(item)) {
    throw new Error("Open a draft or reply compose window before inserting text.");
  }

  await setAsync<void>((callback) =>
    item.body.prependAsync(text, { coercionType: Office.CoercionType.Text }, callback),
  );
}

export async function setBodyText(text: string): Promise<void> {
  const item = getMailboxItem();

  if (!isComposeItem(item)) {
    throw new Error("Open a draft or reply compose window before inserting text.");
  }

  await setAsync<void>((callback) => item.body.setAsync(text, { coercionType: Office.CoercionType.Text }, callback));
}

export async function displayReplyForm(text: string): Promise<void> {
  const item = getMailboxItem();

  if (!item || isComposeItem(item)) {
    throw new Error("Open a received email before drafting a reply.");
  }

  const readItem = item as Office.MessageRead;
  readItem.displayReplyForm(text);
}

export function isComposeMode(): boolean {
  return getSurfaceMode() === "compose";
}

export function getHostInfo(): { host: string; platform: string } {
  return {
    host: String(Office.context?.host ?? "Outlook"),
    platform: String(Office.context?.platform ?? "Office"),
  };
}
