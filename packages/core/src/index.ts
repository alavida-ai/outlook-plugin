export {
  AuthError,
  AuthCacheMissingError,
  AuthCacheCorruptError,
  AuthRefreshFailedError,
  AuthInteractionRequiredError,
  AuthAmbiguousAccountError,
  AuthLockTimeoutError,
} from './auth/errors.js';

export type { TokenCache } from './auth/cache.js';
export { FileTokenCache } from './auth/cache-file.js';
export type { LockOptions } from './auth/cache-file.js';

export {
  buildMsalApp,
  EMBEDDED_CLIENT_ID,
  EMBEDDED_TENANT,
  OUTLOOK_SCOPES,
} from './auth/msal.js';
export type { BuildMsalAppOptions } from './auth/msal.js';

export { resolveAccount } from './auth/accounts.js';

export { getAccessToken } from './auth/silent.js';
export type { AccessTokenResult, GetAccessTokenInput } from './auth/silent.js';

export { loginDeviceCode } from './auth/device-code.js';
export type { LoginDeviceCodeInput, LoginResult } from './auth/device-code.js';

export { logout } from './auth/logout.js';
export type { LogoutInput } from './auth/logout.js';

export { status } from './auth/status.js';
export type { StatusInput } from './auth/status.js';

export { OutlookClient } from './client.js';
export { makeGraphClient } from './graph/client.js';
export type { MakeGraphClientOptions } from './graph/client.js';
export { MeResource } from './resources/me.js';
export type { MeProfile } from './resources/me.js';
export {
  MailResource,
  WELL_KNOWN_FOLDERS,
  INLINE_ATTACHMENT_MAX_BYTES,
  composeLinkFromWebLink,
  normaliseDateForFilter,
  sanitiseAttachmentName,
} from './resources/mail.js';
export type {
  MessageSummary,
  MessageFull,
  MailFolderSummary,
  AttachmentSummary,
  PageEnvelope,
  ListMessagesOptions,
  SearchMessagesOptions,
  ReadMessageOptions,
  DownloadAttachmentResult,
  DraftInput,
  ReplyInput,
  ForwardInput,
  AddAttachmentInput,
  DraftSummary,
  AddAttachmentSummary,
} from './resources/mail.js';
export {
  CalendarResource,
  RECURRENCE_PRESETS,
  ATTENDEE_RESPONSES,
  AVAILABILITY_LEGEND,
  normaliseIso,
  flattenEvent,
} from './resources/calendar.js';
export type {
  EventSummary,
  EventDetail,
  AttendeeSummary,
  RecurrencePreset,
  AttendeeResponse,
  ListEventsOptions,
  CreateEventInput,
  UpdateEventInput,
  RespondInput,
  AvailabilityInput,
  AvailabilityResult,
  AvailabilityScheduleSummary,
} from './resources/calendar.js';
export {
  CoreError,
  NotFoundError,
  ThrottledError,
  ServerError,
  NetworkError,
  liftGraphError,
} from './graph/errors.js';
