unit FilePropInfo;

interface

uses
  Winapi.ActiveX,
  Winapi.PropSys,
  Winapi.PropKey,
  System.TypInfo;

type
  TPropertyItem = record
    Key: PPropertyKey;
    Name: string;
  end;
  PPropertyItem = ^TPropertyItem;

  tagPROPVARIANTHelper = record helper for tagPROPVARIANT
  public
    function TryToString(var str: string): Boolean;
    function ToString: string;
  end;


// PropKey 的分類與系統不一定相同，尚未查到相關資料，僅依屬性內容顯示的歸類

//
// PKEY_PropGroup_General
//
type
  TPropKeyGeneral = (
    PK_FileName,
    PK_ItemFolderNameDisplay,
    PK_ItemFolderPathDisplay,
    PK_ItemNameDisplay,
    PK_ItemType,
    PK_ItemTypeText,
    PK_Size,
    PK_FileAllocationSize,
    PK_DateCreated,
    PK_DateModified,
    PK_DateAccessed,
    PK_FileAttributes
  );
  TPropKeyGeneralFlags = set of TPropKeyGeneral;
const
  PropKeyGeneral: array[TPropKeyGeneral] of TPropertyItem = (
    (Key: @PKEY_FileName;              Name: 'FileName'),
    (Key: @PKEY_ItemFolderNameDisplay; Name: 'ItemFolderNameDisplay'),
    (Key: @PKEY_ItemFolderPathDisplay; Name: 'ItemFolderPathDisplay'),
    (Key: @PKEY_ItemPathDisplay;       Name: 'ItemPathDisplay'),
    (Key: @PKEY_ItemType;              Name: 'ItemType'),
    (Key: @PKEY_ItemTypeText;          Name: 'ItemTypeText'),
    (Key: @PKEY_Size;                  Name: 'Size'),
    (Key: @PKEY_FileAllocationSize;    Name: 'FileAllocationSize'),
    (Key: @PKEY_DateCreated;           Name: 'DateCreated'),
    (Key: @PKEY_DateModified;          Name: 'DateModified'),
    (Key: @PKEY_DateAccessed;          Name: 'DateAccessed'),
    (Key: @PKEY_FileAttributes;        Name: 'FileAttributes')
  );


//
// PKEY_PropGroup_FileSystem
//
type
  TPropKeyFileSystem = (
    PK_AcquisitionID,
    PK_Capacity,
    PK_ComputerName,
    PK_DateAcquired,
    PK_DateArchived,
    PK_DateCompleted,
    PK_DateImported,
    PK_DueDate,
    PK_EndDate,
    PK_FindData,
    PK_FileCount,
    PK_FileDescription,
    PK_FileExtension,
    PK_FileFRN,
    PK_FileOwner,
    PK_FileVersion,
    PK_IsShared,
    PK_ItemPathDisplay,
    PK_OfflineAvailability,
    PK_OfflineStatus,
    PK_OriginalFileName,
    PK_OwnerSID,
    PK_SharedWith,
    PK_ShareUserRating,
    PK_SharingStatus,
    PK_TotalFileSize,
    PK_Trademarks
  );
  TPropKeyFileSystemFlags = set of TPropKeyFileSystem;
const
  PropKeyFileSystem: array[TPropKeyFileSystem] of TPropertyItem = (
    (Key: @PKEY_AcquisitionID;       Name: 'AcquisitionID'),
    (Key: @PKEY_Capacity;            Name: 'Capacity'),
    (Key: @PKEY_ComputerName;        Name: 'ComputerName'),
    (Key: @PKEY_DateAcquired;        Name: 'DateAcquired'),
    (Key: @PKEY_DateArchived;        Name: 'DateArchived'),
    (Key: @PKEY_DateCompleted;       Name: 'DateCompleted'),
    (Key: @PKEY_DateImported;        Name: 'DateImported'),
    (Key: @PKEY_DueDate;             Name: 'DueDate'),
    (Key: @PKEY_EndDate;             Name: 'EndDate'),
    (Key: @PKEY_FindData;            Name: 'FindData'),
    (Key: @PKEY_FileCount;           Name: 'FileCount'),
    (Key: @PKEY_FileDescription;     Name: 'FileDescription'),
    (Key: @PKEY_FileExtension;       Name: 'FileExtension'),
    (Key: @PKEY_FileFRN;             Name: 'FileFRN'),
    (Key: @PKEY_FileOwner;           Name: 'FileOwner'),
    (Key: @PKEY_FileVersion;         Name: 'FileVersion'),
    (Key: @PKEY_IsShared;            Name: 'IsShared'),
    (Key: @PKEY_ItemNameDisplay;     Name: 'ItemNameDisplay'),
    (Key: @PKEY_OfflineAvailability; Name: 'OfflineAvailability'),
    (Key: @PKEY_OfflineStatus;       Name: 'OfflineStatus'),
    (Key: @PKEY_OriginalFileName;    Name: 'OriginalFileName'),
    (Key: @PKEY_OwnerSID;            Name: 'OwnerSID'),
    (Key: @PKEY_SharedWith;          Name: 'SharedWith'),
    (Key: @PKEY_ShareUserRating;     Name: 'ShareUserRating'),
    (Key: @PKEY_SharingStatus;       Name: 'SharingStatus'),
    (Key: @PKEY_TotalFileSize;       Name: 'TotalFileSize'),
    (Key: @PKEY_Trademarks;          Name: 'Trademarks')
  );


//
// PKEY_PropGroup_Description
//
type
  TPropKeyDescription = (
    PK_Title,
    PK_Subject,
    PK_Category,
    PK_Keywords,
    PK_Comment,

    PK_Link_Arguments,
    PK_Link_Comment,
    PK_Link_DateVisited,
    PK_Link_Description,
    PK_Link_Status,
    PK_Link_TargetExtension,
    PK_Link_TargetParsingPath,
    PK_Link_TargetSFGAOFlags,
    PK_Link_TargetSFGAOFlagsStrings,
    PK_Link_TargetUrl
  );
  TPropKeyDescriptionFlags = set of TPropKeyDescription;
const
  PropKeyDescription: array[TPropKeyDescription] of TPropertyItem = (
    (Key: @PKEY_Title;    Name: 'Title'),
    (Key: @PKEY_Subject;  Name: 'Subject'),
    (Key: @PKEY_Category; Name: 'Category'),
    (Key: @PKEY_Keywords; Name: 'Keywords'),
    (Key: @PKEY_Comment;  Name: 'Comment'),

    (Key: @PKEY_Link_Arguments;               Name: 'Arguments'),
    (Key: @PKEY_Link_Comment;                 Name: 'Comment'),
    (Key: @PKEY_Link_DateVisited;             Name: 'DateVisited'),
    (Key: @PKEY_Link_Description;             Name: 'Description'),
    (Key: @PKEY_Link_Status;                  Name: 'Status'),
    (Key: @PKEY_Link_TargetExtension;         Name: 'TargetExtension'),
    (Key: @PKEY_Link_TargetParsingPath;       Name: 'TargetParsingPath'),
    (Key: @PKEY_Link_TargetSFGAOFlags;        Name: 'TargetSFGAOFlags'),
    (Key: @PKEY_Link_TargetSFGAOFlagsStrings; Name: 'TargetSFGAOFlagsStrings'),
    (Key: @PKEY_Link_TargetUrl;               Name: 'TargetUrl')
  );


//
// PKEY_PropGroup_Content
//
type
  TPropKeyContent = (
    PK_ApplicationName,
    PK_Author,
    PK_Company,
    PK_ContainedItems,
    PK_ContentStatus,
    PK_ContentType,
    PK_Copyright,
    PK_FlagColor,
    PK_FlagColorText,
    PK_FlagStatus,
    PK_FlagStatusText,
    PK_FreeSpace,
    PK_FullText,
    PK_Identity,
    PK_Identity_Blob,
    PK_Identity_DisplayName,
    PK_Identity_IsMeIdentity,
    PK_Identity_PrimaryEmailAddress,
    PK_Identity_ProviderID,
    PK_Identity_UniqueID,
    PK_Identity_UserName,
    PK_IdentityProvider_Name,
    PK_IdentityProvider_Picture,
    PK_ImageParsingName,
    PK_Importance,
    PK_ImportanceText,
    PK_IsAttachment,
    PK_IsDefaultNonOwnerSaveLocation,
    PK_IsDefaultSaveLocation,
    PK_IsDeleted,
    PK_IsEncrypted,
    PK_IsFlagged,
    PK_IsFlaggedComplete,
    PK_IsIncomplete,
    PK_IsPinnedToNameSpaceTree,
    PK_IsRead,
    PK_IsSearchOnlyItem,
    PK_IsSendToTarget,
    PK_ItemAuthors,
    PK_ItemClassType,
    PK_ItemDate,
    PK_ItemFolderPathDisplayNarrow,
    PK_ItemName,
    PK_ItemNamePrefix,
    PK_ItemParticipants,
    PK_ItemPathDisplayNarrow,
    PK_ItemUrl,
    PK_Kind,
    PK_KindText,
    PK_Language,
    PK_MileageInformation,
    PK_MIMEType,
    PK_Null,
    PK_ParentalRating,
    PK_ParentalRatingReason,
    PK_ParentalRatingsOrganization,
    PK_ParsingBindContext,
    PK_ParsingName,
    PK_ParsingPath,
    PK_PerceivedType,
    PK_PercentFull,
    PK_Priority,
    PK_PriorityText,
    PK_Project,
    PK_ProviderItemID,
    PK_Rating,
    PK_RatingText,
    PK_Sensitivity,
    PK_SensitivityText,
    PK_SFGAOFlags,
    PK_Shell_OmitFromView,
    PK_SimpleRating,
    PK_SoftwareUsed,
    PK_SourceItem,
    PK_StartDate,
    PK_Status,
    PK_Thumbnail,
    PK_ThumbnailCacheId,
    PK_ThumbnailStream
  );
  TPropKeyContentFlags = set of TPropKeyContent;
const
  PropKeyContent: array[TPropKeyContent] of TPropertyItem = (
    (Key: @PKEY_ApplicationName;               Name: 'ApplicationName'),
    (Key: @PKEY_Author;                        Name: 'Author'),
    (Key: @PKEY_Company;                       Name: 'Company'),
    (Key: @PKEY_ContainedItems;                Name: 'ContainedItems'),
    (Key: @PKEY_ContentStatus;                 Name: 'ContentStatus'),
    (Key: @PKEY_ContentType;                   Name: 'ContentType'),
    (Key: @PKEY_Copyright;                     Name: 'Copyright'),
    (Key: @PKEY_FlagColor;                     Name: 'FlagColor'),
    (Key: @PKEY_FlagColorText;                 Name: 'FlagColorText'),
    (Key: @PKEY_FlagStatus;                    Name: 'FlagStatus'),
    (Key: @PKEY_FlagStatusText;                Name: 'FlagStatusText'),
    (Key: @PKEY_FreeSpace;                     Name: 'FreeSpace'),
    (Key: @PKEY_FullText;                      Name: 'FullText'),
    (Key: @PKEY_Identity;                      Name: 'Identity'),
    (Key: @PKEY_Identity_Blob;                 Name: 'Identity_Blob'),
    (Key: @PKEY_Identity_DisplayName;          Name: 'Identity_DisplayName'),
    (Key: @PKEY_Identity_IsMeIdentity;         Name: 'Identity_IsMeIdentity'),
    (Key: @PKEY_Identity_PrimaryEmailAddress;  Name: 'Identity_PrimaryEmailAddress'),
    (Key: @PKEY_Identity_ProviderID;           Name: 'Identity_ProviderID'),
    (Key: @PKEY_Identity_UniqueID;             Name: 'Identity_UniqueID'),
    (Key: @PKEY_Identity_UserName;             Name: 'Identity_UserName'),
    (Key: @PKEY_IdentityProvider_Name;         Name: 'IdentityProvider_Name'),
    (Key: @PKEY_IdentityProvider_Picture;      Name: 'IdentityProvider_Picture'),
    (Key: @PKEY_ImageParsingName;              Name: 'ImageParsingName'),
    (Key: @PKEY_Importance;                    Name: 'Importance'),
    (Key: @PKEY_ImportanceText;                Name: 'ImportanceText'),
    (Key: @PKEY_IsAttachment;                  Name: 'IsAttachment'),
    (Key: @PKEY_IsDefaultNonOwnerSaveLocation; Name: 'IsDefaultNonOwnerSaveLocation'),
    (Key: @PKEY_IsDefaultSaveLocation;         Name: 'IsDefaultSaveLocation'),
    (Key: @PKEY_IsDeleted;                     Name: 'IsDeleted'),
    (Key: @PKEY_IsEncrypted;                   Name: 'IsEncrypted'),
    (Key: @PKEY_IsFlagged;                     Name: 'IsFlagged'),
    (Key: @PKEY_IsFlaggedComplete;             Name: 'IsFlaggedComplete'),
    (Key: @PKEY_IsIncomplete;                  Name: 'IsIncomplete'),
    (Key: @PKEY_IsPinnedToNameSpaceTree;       Name: 'IsPinnedToNameSpaceTree'),
    (Key: @PKEY_IsRead;                        Name: 'IsRead'),
    (Key: @PKEY_IsSearchOnlyItem;              Name: 'IsSearchOnlyItem'),
    (Key: @PKEY_IsSendToTarget;                Name: 'IsSendToTarget'),
    (Key: @PKEY_ItemAuthors;                   Name: 'ItemAuthors'),
    (Key: @PKEY_ItemClassType;                 Name: 'ItemClassType'),
    (Key: @PKEY_ItemDate;                      Name: 'ItemDate'),
    (Key: @PKEY_ItemFolderPathDisplayNarrow;   Name: 'ItemFolderPathDisplayNarrow'),
    (Key: @PKEY_ItemName;                      Name: 'ItemName'),
    (Key: @PKEY_ItemNamePrefix;                Name: 'ItemNamePrefix'),
    (Key: @PKEY_ItemParticipants;              Name: 'ItemParticipants'),
    (Key: @PKEY_ItemPathDisplayNarrow;         Name: 'ItemPathDisplayNarrow'),
    (Key: @PKEY_ItemUrl;                       Name: 'ItemUrl'),
    (Key: @PKEY_Kind;                          Name: 'Kind'),
    (Key: @PKEY_KindText;                      Name: 'KindText'),
    (Key: @PKEY_Language;                      Name: 'Language'),
    (Key: @PKEY_MileageInformation;            Name: 'MileageInformation'),
    (Key: @PKEY_MIMEType;                      Name: 'MIMEType'),
    (Key: @PKEY_Null;                          Name: 'Null'),
    (Key: @PKEY_ParentalRating;                Name: 'ParentalRating'),
    (Key: @PKEY_ParentalRatingReason;          Name: 'ParentalRatingReason'),
    (Key: @PKEY_ParentalRatingsOrganization;   Name: 'ParentalRatingsOrganization'),
    (Key: @PKEY_ParsingBindContext;            Name: 'ParsingBindContext'),
    (Key: @PKEY_ParsingName;                   Name: 'ParsingName'),
    (Key: @PKEY_ParsingPath;                   Name: 'ParsingPath'),
    (Key: @PKEY_PerceivedType;                 Name: 'PerceivedType'),
    (Key: @PKEY_PercentFull;                   Name: 'PercentFull'),
    (Key: @PKEY_Priority;                      Name: 'Priority'),
    (Key: @PKEY_PriorityText;                  Name: 'PriorityText'),
    (Key: @PKEY_Project;                       Name: 'Project'),
    (Key: @PKEY_ProviderItemID;                Name: 'ProviderItemID'),
    (Key: @PKEY_Rating;                        Name: 'Rating'),
    (Key: @PKEY_RatingText;                    Name: 'RatingText'),
    (Key: @PKEY_Sensitivity;                   Name: 'Sensitivity'),
    (Key: @PKEY_SensitivityText;               Name: 'SensitivityText'),
    (Key: @PKEY_SFGAOFlags;                    Name: 'SFGAOFlags'),
    (Key: @PKEY_Shell_OmitFromView;            Name: 'Shell_OmitFromView'),
    (Key: @PKEY_SimpleRating;                  Name: 'SimpleRating'),
    (Key: @PKEY_SoftwareUsed;                  Name: 'SoftwareUsed'),
    (Key: @PKEY_SourceItem;                    Name: 'SourceItem'),
    (Key: @PKEY_StartDate;                     Name: 'StartDate'),
    (Key: @PKEY_Status;                        Name: 'Status'),
    (Key: @PKEY_Thumbnail;                     Name: 'Thumbnail'),
    (Key: @PKEY_ThumbnailCacheId;              Name: 'ThumbnailCacheId'),
    (Key: @PKEY_ThumbnailStream;               Name: 'ThumbnailStream')
  );


//
// PKEY_PropGroup_Audio
//
type
  TPropKeyAudio = (
    PK_Audio_Format,
    PK_Audio_EncodingBitrate,
    PK_Audio_SampleRate,
    PK_Audio_SampleSize,
    PK_Audio_ChannelCount,
    PK_Audio_StreamNumber,
    PK_Audio_StreamName,
    PK_Audio_Compression,
    PK_Audio_IsVariableBitRate,
    PK_Audio_PeakValue
  );
  TPropKeyAudioFlags = set of TPropKeyAudio;
const
  PropKeyAudio: array[TPropKeyAudio] of TPropertyItem = (
    (Key: @PKEY_Audio_Format;            Name: 'Format'),
    (Key: @PKEY_Audio_EncodingBitrate;   Name: 'EncodingBitrate'),
    (Key: @PKEY_Audio_SampleRate;        Name: 'SampleRate'),
    (Key: @PKEY_Audio_SampleSize;        Name: 'SampleSize'),
    (Key: @PKEY_Audio_ChannelCount;      Name: 'ChannelCount'),
    (Key: @PKEY_Audio_StreamNumber;      Name: 'StreamNumber'),
    (Key: @PKEY_Audio_StreamName;        Name: 'StreamName'),
    (Key: @PKEY_Audio_Compression;       Name: 'Compression'),
    (Key: @PKEY_Audio_IsVariableBitRate; Name: 'IsVariableBitRate'),
    (Key: @PKEY_Audio_PeakValue;         Name: 'PeakValue')
  );


//
// PKEY_PropGroup_Calendar
//
type
  TPropKeyCalendar = (
    PK_Calendar_Duration,
    PK_Calendar_IsOnline,
    PK_Calendar_IsRecurring,
    PK_Calendar_Location,
    PK_Calendar_OptionalAttendeeAddresses,
    PK_Calendar_OptionalAttendeeNames,
    PK_Calendar_OrganizerAddress,
    PK_Calendar_OrganizerName,
    PK_Calendar_ReminderTime,
    PK_Calendar_RequiredAttendeeAddresses,
    PK_Calendar_RequiredAttendeeNames,
    PK_Calendar_Resources,
    PK_Calendar_ResponseStatus,
    PK_Calendar_ShowTimeAs,
    PK_Calendar_ShowTimeAsText
  );
  TPropKeyCalendarFlags = set of TPropKeyCalendar;
const
  PropKeyCalendar: array[TPropKeyCalendar] of TPropertyItem = (
    (Key: @PKEY_Calendar_Duration;                  Name: 'Duration'),
    (Key: @PKEY_Calendar_IsOnline;                  Name: 'IsOnline'),
    (Key: @PKEY_Calendar_IsRecurring;               Name: 'IsRecurring'),
    (Key: @PKEY_Calendar_Location;                  Name: 'Location'),
    (Key: @PKEY_Calendar_OptionalAttendeeAddresses; Name: 'OptionalAttendeeAddresses'),
    (Key: @PKEY_Calendar_OptionalAttendeeNames;     Name: 'OptionalAttendeeNames'),
    (Key: @PKEY_Calendar_OrganizerAddress;          Name: 'OrganizerAddress'),
    (Key: @PKEY_Calendar_OrganizerName;             Name: 'OrganizerName'),
    (Key: @PKEY_Calendar_ReminderTime;              Name: 'ReminderTime'),
    (Key: @PKEY_Calendar_RequiredAttendeeAddresses; Name: 'RequiredAttendeeAddresses'),
    (Key: @PKEY_Calendar_RequiredAttendeeNames;     Name: 'RequiredAttendeeNames'),
    (Key: @PKEY_Calendar_Resources;                 Name: 'Resources'),
    (Key: @PKEY_Calendar_ResponseStatus;            Name: 'ResponseStatus'),
    (Key: @PKEY_Calendar_ShowTimeAs;                Name: 'ShowTimeAs'),
    (Key: @PKEY_Calendar_ShowTimeAsText;            Name: 'ShowTimeAsText')
  );

// PKEY_PropGroup_Camera
type
  TPropKeyCamera = (
    PK_Photo_Aperture,
    PK_Photo_ApertureDenominator,
    PK_Photo_ApertureNumerator,
    PK_Photo_CameraManufacturer,
    PK_Photo_CameraModel,
    PK_Photo_CameraSerialNumber,
    PK_Photo_DateTaken,
    PK_Photo_DigitalZoom,
    PK_Photo_DigitalZoomDenominator,
    PK_Photo_DigitalZoomNumerator,
    PK_Photo_Event,
    PK_Photo_ExposureBias,
    PK_Photo_ExposureBiasDenominator,
    PK_Photo_ExposureBiasNumerator,
    PK_Photo_ExposureIndex,
    PK_Photo_ExposureIndexDenominator,
    PK_Photo_ExposureIndexNumerator,
    PK_Photo_ExposureProgram,
    PK_Photo_ExposureProgramText,
    PK_Photo_ExposureTime,
    PK_Photo_ExposureTimeDenominator,
    PK_Photo_ExposureTimeNumerator,
    PK_Photo_Flash,
    PK_Photo_FlashEnergy,
    PK_Photo_FlashEnergyDenominator,
    PK_Photo_FlashEnergyNumerator,
    PK_Photo_FlashManufacturer,
    PK_Photo_FlashModel,
    PK_Photo_FlashText,
    PK_Photo_FNumber,
    PK_Photo_FNumberDenominator,
    PK_Photo_FNumberNumerator,
    PK_Photo_FocalLength,
    PK_Photo_FocalLengthDenominator,
    PK_Photo_FocalLengthInFilm,
    PK_Photo_FocalLengthNumerator,
    PK_Photo_FocalPlaneXResolution,
    PK_Photo_FocalPlaneXResolutionDenominator,
    PK_Photo_FocalPlaneXResolutionNumerator,
    PK_Photo_FocalPlaneYResolution,
    PK_Photo_FocalPlaneYResolutionDenominator,
    PK_Photo_FocalPlaneYResolutionNumerator,
    PK_Photo_ISOSpeed,
    PK_Photo_LensManufacturer,
    PK_Photo_LensModel,
    PK_Photo_LightSource,
    PK_Photo_MeteringMode,
    PK_Photo_MeteringModeText,
    PK_Photo_MaxAperture,
    PK_Photo_MaxApertureDenominator,
    PK_Photo_MaxApertureNumerator,
    PK_Photo_Orientation,
    PK_Photo_OrientationText,
    PK_Photo_ShutterSpeed,
    PK_Photo_ShutterSpeedDenominator,
    PK_Photo_ShutterSpeedNumerator,
    PK_Photo_SubjectDistance,
    PK_Photo_SubjectDistanceDenominator,
    PK_Photo_SubjectDistanceNumerator
  );
  TPropKeyCameraFlags = set of TPropKeyCamera;
const
  PropKeyCamera: array[TPropKeyCamera] of TPropertyItem = (
    (Key: @PKEY_Photo_Aperture;                 Name: 'Aperture'),
    (Key: @PKEY_Photo_ApertureDenominator;      Name: 'ApertureDenominator'),
    (Key: @PKEY_Photo_ApertureNumerator;        Name: 'ApertureNumerator'),
    (Key: @PKEY_Photo_CameraManufacturer;       Name: 'CameraManufacturer'),
    (Key: @PKEY_Photo_CameraModel;              Name: 'CameraModel'),
    (Key: @PKEY_Photo_CameraSerialNumber;       Name: 'CameraSerialNumber'),
    (Key: @PKEY_Photo_DateTaken;                Name: 'DateTaken'),
    (Key: @PKEY_Photo_DigitalZoom;              Name: 'DigitalZoom'),
    (Key: @PKEY_Photo_DigitalZoomDenominator;   Name: 'DigitalZoomDenominator'),
    (Key: @PKEY_Photo_DigitalZoomNumerator;     Name: 'DigitalZoomNumerator'),
    (Key: @PKEY_Photo_Event;                    Name: 'Event'),
    (Key: @PKEY_Photo_ExposureBias;             Name: 'ExposureBias'),
    (Key: @PKEY_Photo_ExposureBiasDenominator;  Name: 'ExposureBiasDenominator'),
    (Key: @PKEY_Photo_ExposureBiasNumerator;    Name: 'ExposureBiasNumerator'),
    (Key: @PKEY_Photo_ExposureIndex;            Name: 'ExposureIndex'),
    (Key: @PKEY_Photo_ExposureIndexDenominator; Name: 'ExposureIndexDenominator'),
    (Key: @PKEY_Photo_ExposureIndexNumerator;   Name: 'ExposureIndexNumerator'),
    (Key: @PKEY_Photo_ExposureProgram;          Name: 'ExposureProgram'),
    (Key: @PKEY_Photo_ExposureProgramText;      Name: 'ExposureProgramText'),
    (Key: @PKEY_Photo_ExposureTime;             Name: 'ExposureTime'),
    (Key: @PKEY_Photo_ExposureTimeDenominator;  Name: 'ExposureTimeDenominator'),
    (Key: @PKEY_Photo_ExposureTimeNumerator;    Name: 'ExposureTimeNumerator'),
    (Key: @PKEY_Photo_Flash;                    Name: 'Flash'),
    (Key: @PKEY_Photo_FlashEnergy;              Name: 'FlashEnergy'),
    (Key: @PKEY_Photo_FlashEnergyDenominator;   Name: 'FlashEnergyDenominator'),
    (Key: @PKEY_Photo_FlashEnergyNumerator;     Name: 'FlashEnergyNumerator'),
    (Key: @PKEY_Photo_FlashManufacturer;        Name: 'FlashManufacturer'),
    (Key: @PKEY_Photo_FlashModel;               Name: 'FlashModel'),
    (Key: @PKEY_Photo_FlashText;                Name: 'FlashText'),
    (Key: @PKEY_Photo_FNumber;                  Name: 'FNumber'),
    (Key: @PKEY_Photo_FNumberDenominator;       Name: 'FNumberDenominator'),
    (Key: @PKEY_Photo_FNumberNumerator;         Name: 'FNumberNumerator'),
    (Key: @PKEY_Photo_FocalLength;              Name: 'FocalLength'),
    (Key: @PKEY_Photo_FocalLengthDenominator;   Name: 'FocalLengthDenominator'),
    (Key: @PKEY_Photo_FocalLengthInFilm;        Name: 'FocalLengthInFilm'),
    (Key: @PKEY_Photo_FocalLengthNumerator;     Name: 'FocalLengthNumerator'),
    (Key: @PKEY_Photo_FocalPlaneXResolution;            Name: 'FocalPlaneXResolution'),
    (Key: @PKEY_Photo_FocalPlaneXResolutionDenominator; Name: 'FocalPlaneXResolutionDenominator'),
    (Key: @PKEY_Photo_FocalPlaneXResolutionNumerator;   Name: 'FocalPlaneXResolutionNumerator'),
    (Key: @PKEY_Photo_FocalPlaneYResolution;            Name: 'FocalPlaneYResolution'),
    (Key: @PKEY_Photo_FocalPlaneYResolutionDenominator; Name: 'FocalPlaneYResolutionDenominator'),
    (Key: @PKEY_Photo_FocalPlaneYResolutionNumerator;   Name: 'FocalPlaneYResolutionNumerator'),
    (Key: @PKEY_Photo_ISOSpeed;                   Name: 'ISOSpeed'),
    (Key: @PKEY_Photo_LensManufacturer;           Name: 'LensManufacturer'),
    (Key: @PKEY_Photo_LensModel;                  Name: 'LensModel'),
    (Key: @PKEY_Photo_LightSource;                Name: 'LightSource'),
    (Key: @PKEY_Photo_MeteringMode;               Name: 'MeteringMode'),
    (Key: @PKEY_Photo_MeteringModeText;           Name: 'MeteringModeText'),
    (Key: @PKEY_Photo_MaxAperture;                Name: 'MaxAperture'),
    (Key: @PKEY_Photo_MaxApertureDenominator;     Name: 'MaxApertureDenominator'),
    (Key: @PKEY_Photo_MaxApertureNumerator;       Name: 'MaxApertureNumerator'),
    (Key: @PKEY_Photo_Orientation;                Name: 'Orientation'),
    (Key: @PKEY_Photo_OrientationText;            Name: 'OrientationText'),
    (Key: @PKEY_Photo_ShutterSpeed;               Name: 'ShutterSpeed'),
    (Key: @PKEY_Photo_ShutterSpeedDenominator;    Name: 'ShutterSpeedDenominator'),
    (Key: @PKEY_Photo_ShutterSpeedNumerator;      Name: 'ShutterSpeedNumerator'),
    (Key: @PKEY_Photo_SubjectDistance;            Name: 'SubjectDistance'),
    (Key: @PKEY_Photo_SubjectDistanceDenominator; Name: 'SubjectDistanceDenominator'),
    (Key: @PKEY_Photo_SubjectDistanceNumerator;   Name: 'SubjectDistanceNumerator')
  );

//
// PKEY_PropGroup_Contact
//
type
  TPropKeyContact = (
    PK_Contact_Anniversary,
    PK_Contact_AssistantName,
    PK_Contact_AssistantTelephone,
    PK_Contact_Birthday,
    PK_Contact_BusinessAddress,
    PK_Contact_BusinessAddressCity,
    PK_Contact_BusinessAddressCountry,
    PK_Contact_BusinessAddressPostalCode,
    PK_Contact_BusinessAddressPostOfficeBox,
    PK_Contact_BusinessAddressState,
    PK_Contact_BusinessAddressStreet,
    PK_Contact_BusinessFaxNumber,
    PK_Contact_BusinessHomePage,
    PK_Contact_BusinessTelephone,
    PK_Contact_CallbackTelephone,
    PK_Contact_CarTelephone,
    PK_Contact_Children,
    PK_Contact_CompanyMainTelephone,
    PK_Contact_Department,
    PK_Contact_EmailAddress,
    PK_Contact_EmailAddress2,
    PK_Contact_EmailAddress3,
    PK_Contact_EmailAddresses,
    PK_Contact_EmailName,
    PK_Contact_FileAsName,
    PK_Contact_FirstName,
    PK_Contact_FullName,
    PK_Contact_Gender,
    PK_Contact_GenderValue,
    PK_Contact_Hobbies,
    PK_Contact_HomeAddress,
    PK_Contact_HomeAddressCity,
    PK_Contact_HomeAddressCountry,
    PK_Contact_HomeAddressPostalCode,
    PK_Contact_HomeAddressPostOfficeBox,
    PK_Contact_HomeAddressState,
    PK_Contact_HomeAddressStreet,
    PK_Contact_HomeFaxNumber,
    PK_Contact_HomeTelephone,
    PK_Contact_IMAddress,
    PK_Contact_Initials,
    PK_Contact_JA_CompanyNamePhonetic,
    PK_Contact_JA_FirstNamePhonetic,
    PK_Contact_JA_LastNamePhonetic,
    PK_Contact_JobTitle,
    PK_Contact_Label,
    PK_Contact_LastName,
    PK_Contact_MailingAddress,
    PK_Contact_MiddleName,
    PK_Contact_MobileTelephone,
    PK_Contact_NickName,
    PK_Contact_OfficeLocation,
    PK_Contact_OtherAddress,
    PK_Contact_OtherAddressCity,
    PK_Contact_OtherAddressCountry,
    PK_Contact_OtherAddressPostalCode,
    PK_Contact_OtherAddressPostOfficeBox,
    PK_Contact_OtherAddressState,
    PK_Contact_OtherAddressStreet,
    PK_Contact_PagerTelephone,
    PK_Contact_PersonalTitle,
    PK_Contact_PrimaryAddressCity,
    PK_Contact_PrimaryAddressCountry,
    PK_Contact_PrimaryAddressPostalCode,
    PK_Contact_PrimaryAddressPostOfficeBox,
    PK_Contact_PrimaryAddressState,
    PK_Contact_PrimaryAddressStreet,
    PK_Contact_PrimaryEmailAddress,
    PK_Contact_PrimaryTelephone,
    PK_Contact_Profession,
    PK_Contact_SpouseName,
    PK_Contact_Suffix,
    PK_Contact_TelexNumber,
    PK_Contact_TTYTDDTelephone,
    PK_Contact_WebPage
  );
  TPropKeyContactFlags = set of TPropKeyContact;
const
  PropKeyContact: array[TPropKeyContact] of TPropertyItem = (
    (Key: @PKEY_Contact_Anniversary;                  Name: 'Anniversary'),
    (Key: @PKEY_Contact_AssistantName;                Name: 'AssistantName'),
    (Key: @PKEY_Contact_AssistantTelephone;           Name: 'AssistantTelephone'),
    (Key: @PKEY_Contact_Birthday;                     Name: 'Birthday'),
    (Key: @PKEY_Contact_BusinessAddress;              Name: 'BusinessAddress'),
    (Key: @PKEY_Contact_BusinessAddressCity;          Name: 'BusinessAddressCity'),
    (Key: @PKEY_Contact_BusinessAddressCountry;       Name: 'BusinessAddressCountry'),
    (Key: @PKEY_Contact_BusinessAddressPostalCode;    Name: 'BusinessAddressPostalCode'),
    (Key: @PKEY_Contact_BusinessAddressPostOfficeBox; Name: 'BusinessAddressPostOfficeBox'),
    (Key: @PKEY_Contact_BusinessAddressState;         Name: 'BusinessAddressState'),
    (Key: @PKEY_Contact_BusinessAddressStreet;        Name: 'BusinessAddressStreet'),
    (Key: @PKEY_Contact_BusinessFaxNumber;            Name: 'BusinessFaxNumber'),
    (Key: @PKEY_Contact_BusinessHomePage;             Name: 'BusinessHomePage'),
    (Key: @PKEY_Contact_BusinessTelephone;            Name: 'BusinessTelephone'),
    (Key: @PKEY_Contact_CallbackTelephone;            Name: 'CallbackTelephone'),
    (Key: @PKEY_Contact_CarTelephone;                 Name: 'CarTelephone'),
    (Key: @PKEY_Contact_Children;                     Name: 'Children'),
    (Key: @PKEY_Contact_CompanyMainTelephone;         Name: 'CompanyMainTelephone'),
    (Key: @PKEY_Contact_Department;                   Name: 'Department'),
    (Key: @PKEY_Contact_EmailAddress;                 Name: 'EmailAddress'),
    (Key: @PKEY_Contact_EmailAddress2;                Name: 'EmailAddress2'),
    (Key: @PKEY_Contact_EmailAddress3;                Name: 'EmailAddress3'),
    (Key: @PKEY_Contact_EmailAddresses;               Name: 'EmailAddresses'),
    (Key: @PKEY_Contact_EmailName;                    Name: 'EmailName'),
    (Key: @PKEY_Contact_FileAsName;                   Name: 'FileAsName'),
    (Key: @PKEY_Contact_FirstName;                    Name: 'FirstName'),
    (Key: @PKEY_Contact_FullName;                     Name: 'FullName'),
    (Key: @PKEY_Contact_Gender;                       Name: 'Gender'),
    (Key: @PKEY_Contact_GenderValue;                  Name: 'GenderValue'),
    (Key: @PKEY_Contact_Hobbies;                      Name: 'Hobbies'),
    (Key: @PKEY_Contact_HomeAddress;                  Name: 'HomeAddress'),
    (Key: @PKEY_Contact_HomeAddressCity;              Name: 'HomeAddressCity'),
    (Key: @PKEY_Contact_HomeAddressCountry;           Name: 'HomeAddressCountry'),
    (Key: @PKEY_Contact_HomeAddressPostalCode;        Name: 'HomeAddressPostalCode'),
    (Key: @PKEY_Contact_HomeAddressPostOfficeBox;     Name: 'HomeAddressPostOfficeBox'),
    (Key: @PKEY_Contact_HomeAddressState;             Name: 'HomeAddressState'),
    (Key: @PKEY_Contact_HomeAddressStreet;            Name: 'HomeAddressStreet'),
    (Key: @PKEY_Contact_HomeFaxNumber;                Name: 'HomeFaxNumber'),
    (Key: @PKEY_Contact_HomeTelephone;                Name: 'HomeTelephone'),
    (Key: @PKEY_Contact_IMAddress;                    Name: 'IMAddress'),
    (Key: @PKEY_Contact_Initials;                     Name: 'Initials'),
    (Key: @PKEY_Contact_JA_CompanyNamePhonetic;       Name: 'JA_CompanyNamePhonetic'),
    (Key: @PKEY_Contact_JA_FirstNamePhonetic;         Name: 'JA_FirstNamePhonetic'),
    (Key: @PKEY_Contact_JA_LastNamePhonetic;          Name: 'JA_LastNamePhonetic'),
    (Key: @PKEY_Contact_JobTitle;                     Name: 'JobTitle'),
    (Key: @PKEY_Contact_Label;                        Name: 'Label'),
    (Key: @PKEY_Contact_LastName;                     Name: 'LastName'),
    (Key: @PKEY_Contact_MailingAddress;               Name: 'MailingAddress'),
    (Key: @PKEY_Contact_MiddleName;                   Name: 'MiddleName'),
    (Key: @PKEY_Contact_MobileTelephone;              Name: 'MobileTelephone'),
    (Key: @PKEY_Contact_NickName;                     Name: 'NickName'),
    (Key: @PKEY_Contact_OfficeLocation;               Name: 'OfficeLocation'),
    (Key: @PKEY_Contact_OtherAddress;                 Name: 'OtherAddress'),
    (Key: @PKEY_Contact_OtherAddressCity;             Name: 'OtherAddressCity'),
    (Key: @PKEY_Contact_OtherAddressCountry;          Name: 'OtherAddressCountry'),
    (Key: @PKEY_Contact_OtherAddressPostalCode;       Name: 'OtherAddressPostalCode'),
    (Key: @PKEY_Contact_OtherAddressPostOfficeBox;    Name: 'OtherAddressPostOfficeBox'),
    (Key: @PKEY_Contact_OtherAddressState;            Name: 'OtherAddressState'),
    (Key: @PKEY_Contact_OtherAddressStreet;           Name: 'OtherAddressStreet'),
    (Key: @PKEY_Contact_PagerTelephone;               Name: 'PagerTelephone'),
    (Key: @PKEY_Contact_PersonalTitle;                Name: 'PersonalTitle'),
    (Key: @PKEY_Contact_PrimaryAddressCity;           Name: 'PrimaryAddressCity'),
    (Key: @PKEY_Contact_PrimaryAddressCountry;        Name: 'PrimaryAddressCountry'),
    (Key: @PKEY_Contact_PrimaryAddressPostalCode;     Name: 'PrimaryAddressPostalCode'),
    (Key: @PKEY_Contact_PrimaryAddressPostOfficeBox;  Name: 'PrimaryAddressPostOfficeBox'),
    (Key: @PKEY_Contact_PrimaryAddressState;          Name: 'PrimaryAddressState'),
    (Key: @PKEY_Contact_PrimaryAddressStreet;         Name: 'PrimaryAddressStreet'),
    (Key: @PKEY_Contact_PrimaryEmailAddress;          Name: 'PrimaryEmailAddress'),
    (Key: @PKEY_Contact_PrimaryTelephone;             Name: 'PrimaryTelephone'),
    (Key: @PKEY_Contact_Profession;                   Name: 'Profession'),
    (Key: @PKEY_Contact_SpouseName;                   Name: 'SpouseName'),
    (Key: @PKEY_Contact_Suffix;                       Name: 'Suffix'),
    (Key: @PKEY_Contact_TelexNumber;                  Name: 'TelexNumber'),
    (Key: @PKEY_Contact_TTYTDDTelephone;              Name: 'TTYTDDTelephone'),
    (Key: @PKEY_Contact_WebPage;                      Name: 'WebPage')
  );


//
// PKEY_PropGroup_GPS
//
type
  TPropKeyGPS = (
    PK_GPS_Altitude,
    PK_GPS_AltitudeDenominator,
    PK_GPS_AltitudeNumerator,
    PK_GPS_AltitudeRef,
    PK_GPS_AreaInformation,
    PK_GPS_Date,
    PK_GPS_DestBearing,
    PK_GPS_DestBearingDenominator,
    PK_GPS_DestBearingNumerator,
    PK_GPS_DestBearingRef,
    PK_GPS_DestDistance,
    PK_GPS_DestDistanceDenominator,
    PK_GPS_DestDistanceNumerator,
    PK_GPS_DestDistanceRef,
    PK_GPS_DestLatitude,
    PK_GPS_DestLatitudeDenominator,
    PK_GPS_DestLatitudeNumerator,
    PK_GPS_DestLatitudeRef,
    PK_GPS_DestLongitude,
    PK_GPS_DestLongitudeDenominator,
    PK_GPS_DestLongitudeNumerator,
    PK_GPS_DestLongitudeRef,
    PK_GPS_Differential,
    PK_GPS_DOP,
    PK_GPS_DOPDenominator,
    PK_GPS_DOPNumerator,
    PK_GPS_ImgDirection,
    PK_GPS_ImgDirectionDenominator,
    PK_GPS_ImgDirectionNumerator,
    PK_GPS_ImgDirectionRef,
    PK_GPS_Latitude,
    PK_GPS_LatitudeDenominator,
    PK_GPS_LatitudeNumerator,
    PK_GPS_LatitudeRef,
    PK_GPS_Longitude,
    PK_GPS_LongitudeDenominator,
    PK_GPS_LongitudeNumerator,
    PK_GPS_LongitudeRef,
    PK_GPS_MapDatum,
    PK_GPS_MeasureMode,
    PK_GPS_ProcessingMethod,
    PK_GPS_Satellites,
    PK_GPS_Speed,
    PK_GPS_SpeedDenominator,
    PK_GPS_SpeedNumerator,
    PK_GPS_SpeedRef,
    PK_GPS_Status,
    PK_GPS_Track,
    PK_GPS_TrackDenominator,
    PK_GPS_TrackNumerator,
    PK_GPS_TrackRef,
    PK_GPS_VersionID
  );
  TPropKeyGPSFlags = set of TPropKeyGPS;
const
  PropKeyGPS: array[TPropKeyGPS] of TPropertyItem = (
    (Key: @PKEY_GPS_Altitude;                 Name: 'Altitude'),
    (Key: @PKEY_GPS_AltitudeDenominator;      Name: 'AltitudeDenominator'),
    (Key: @PKEY_GPS_AltitudeNumerator;        Name: 'AltitudeNumerator'),
    (Key: @PKEY_GPS_AltitudeRef;              Name: 'AltitudeRef'),
    (Key: @PKEY_GPS_AreaInformation;          Name: 'AreaInformation'),
    (Key: @PKEY_GPS_Date;                     Name: 'Date'),
    (Key: @PKEY_GPS_DestBearing;              Name: 'DestBearing'),
    (Key: @PKEY_GPS_DestBearingDenominator;   Name: 'DestBearingDenominator'),
    (Key: @PKEY_GPS_DestBearingNumerator;     Name: 'DestBearingNumerator'),
    (Key: @PKEY_GPS_DestBearingRef;           Name: 'DestBearingRef'),
    (Key: @PKEY_GPS_DestDistance;             Name: 'DestDistance'),
    (Key: @PKEY_GPS_DestDistanceDenominator;  Name: 'DestDistanceDenominator'),
    (Key: @PKEY_GPS_DestDistanceNumerator;    Name: 'DestDistanceNumerator'),
    (Key: @PKEY_GPS_DestDistanceRef;          Name: 'DestDistanceRef'),
    (Key: @PKEY_GPS_DestLatitude;             Name: 'DestLatitude'),
    (Key: @PKEY_GPS_DestLatitudeDenominator;  Name: 'DestLatitudeDenominator'),
    (Key: @PKEY_GPS_DestLatitudeNumerator;    Name: 'DestLatitudeNumerator'),
    (Key: @PKEY_GPS_DestLatitudeRef;          Name: 'DestLatitudeRef'),
    (Key: @PKEY_GPS_DestLongitude;            Name: 'DestLongitude'),
    (Key: @PKEY_GPS_DestLongitudeDenominator; Name: 'DestLongitudeDenominator'),
    (Key: @PKEY_GPS_DestLongitudeNumerator;   Name: 'DestLongitudeNumerator'),
    (Key: @PKEY_GPS_DestLongitudeRef;         Name: 'DestLongitudeRef'),
    (Key: @PKEY_GPS_Differential;             Name: 'Differential'),
    (Key: @PKEY_GPS_DOP;                      Name: 'DOP'),
    (Key: @PKEY_GPS_DOPDenominator;           Name: 'DOPDenominator'),
    (Key: @PKEY_GPS_DOPNumerator;             Name: 'DOPNumerator'),
    (Key: @PKEY_GPS_ImgDirection;             Name: 'ImgDirection'),
    (Key: @PKEY_GPS_ImgDirectionDenominator;  Name: 'ImgDirectionDenominator'),
    (Key: @PKEY_GPS_ImgDirectionNumerator;    Name: 'ImgDirectionNumerator'),
    (Key: @PKEY_GPS_ImgDirectionRef;          Name: 'ImgDirectionRef'),
    (Key: @PKEY_GPS_Latitude;                 Name: 'Latitude'),
    (Key: @PKEY_GPS_LatitudeDenominator;      Name: 'LatitudeDenominator'),
    (Key: @PKEY_GPS_LatitudeNumerator;        Name: 'LatitudeNumerator'),
    (Key: @PKEY_GPS_LatitudeRef;              Name: 'LatitudeRef'),
    (Key: @PKEY_GPS_Longitude;                Name: 'Longitude'),
    (Key: @PKEY_GPS_LongitudeDenominator;     Name: 'LongitudeDenominator'),
    (Key: @PKEY_GPS_LongitudeNumerator;       Name: 'LongitudeNumerator'),
    (Key: @PKEY_GPS_LongitudeRef;             Name: 'LongitudeRef'),
    (Key: @PKEY_GPS_MapDatum;                 Name: 'MapDatum'),
    (Key: @PKEY_GPS_MeasureMode;              Name: 'MeasureMode'),
    (Key: @PKEY_GPS_ProcessingMethod;         Name: 'ProcessingMethod'),
    (Key: @PKEY_GPS_Satellites;               Name: 'Satellites'),
    (Key: @PKEY_GPS_Speed;                    Name: 'Speed'),
    (Key: @PKEY_GPS_SpeedDenominator;         Name: 'SpeedDenominator'),
    (Key: @PKEY_GPS_SpeedNumerator;           Name: 'SpeedNumerator'),
    (Key: @PKEY_GPS_SpeedRef;                 Name: 'SpeedRef'),
    (Key: @PKEY_GPS_Status;                   Name: 'Status'),
    (Key: @PKEY_GPS_Track;                    Name: 'Track'),
    (Key: @PKEY_GPS_TrackDenominator;         Name: 'TrackDenominator'),
    (Key: @PKEY_GPS_TrackNumerator;           Name: 'TrackNumerator'),
    (Key: @PKEY_GPS_TrackRef;                 Name: 'TrackRef'),
    (Key: @PKEY_GPS_VersionID;                Name: 'VersionID')
  );


//
// PKEY_PropGroup_Image
//
type
  TPropKeyImage = (
    PK_Image_BitDepth,
    PK_Image_ColorSpace,
    PK_Image_CompressedBitsPerPixel,
    PK_Image_CompressedBitsPerPixelDenominator,
    PK_Image_CompressedBitsPerPixelNumerator,
    PK_Image_Compression,
    PK_Image_CompressionText,
    PK_Image_Dimensions,
    PK_Image_HorizontalResolution,
    PK_Image_HorizontalSize,
    PK_Image_ImageID,
    PK_Image_ResolutionUnit,
    PK_Image_VerticalResolution,
    PK_Image_VerticalSize
  );
  TPropKeyImageFlags = set of TPropKeyImage;
const
  PropKeyImage: array[TPropKeyImage] of TPropertyItem = (
    (Key: @PKEY_Image_BitDepth;             Name: 'BitDepth'),
    (Key: @PKEY_Image_ColorSpace;           Name: 'ColorSpace'),
    (Key: @PKEY_Image_CompressedBitsPerPixel;            Name: 'CompressedBitsPerPixel'),
    (Key: @PKEY_Image_CompressedBitsPerPixelDenominator; Name: 'CompressedBitsPerPixelDenominator'),
    (Key: @PKEY_Image_CompressedBitsPerPixelNumerator;   Name: 'CompressedBitsPerPixelNumerator'),
    (Key: @PKEY_Image_Compression;          Name: 'Compression'),
    (Key: @PKEY_Image_CompressionText;      Name: 'CompressionText'),
    (Key: @PKEY_Image_Dimensions;           Name: 'Dimensions'),
    (Key: @PKEY_Image_HorizontalResolution; Name: 'HorizontalResolution'),
    (Key: @PKEY_Image_HorizontalSize;       Name: 'HorizontalSize'),
    (Key: @PKEY_Image_ImageID;              Name: 'ImageID'),
    (Key: @PKEY_Image_ResolutionUnit;       Name: 'ResolutionUnit'),
    (Key: @PKEY_Image_VerticalResolution;   Name: 'VerticalResolution'),
    (Key: @PKEY_Image_VerticalSize;         Name: 'VerticalSize')
  );


//
// PKEY_PropGroup_Media
//
type
  TPropKeyMedia = (
    PK_Media_Duration,
    PK_Media_ClassPrimaryID,
    PK_Media_ClassSecondaryID,
    PK_Media_DVDID,
    PK_Media_MCDI,
    PK_Media_MetadataContentProvider,
    PK_Media_ContentDistributor,
    PK_Media_Producer,
    PK_Media_Writer,
    PK_Media_CollectionGroupID,
    PK_Media_CollectionID,
    PK_Media_ContentID,
    PK_Media_CreatorApplication,
    PK_Media_CreatorApplicationVersion,
    PK_Media_Publisher,
    PK_Media_AuthorUrl,
    PK_Media_PromotionUrl,
    PK_Media_UserWebUrl,
    PK_Media_UniqueFileIdentifier,
    PK_Media_EncodedBy,
    PK_Media_EncodingSettings,
    PK_Media_ProviderRating,
    PK_Media_ProtectionType,
    PK_Media_ProviderStyle,
    PK_Media_UserNoAutoInfo,

    PK_Media_AverageLevel,
    PK_Media_DateEncoded,
    PK_Media_DateReleased,
    PK_Media_Year,
    PK_Media_SubTitle,
    PK_Media_FrameCount,
    PK_Media_SubscriptionContentId
  );
  TPropKeyMediaFlags = set of TPropKeyMedia;
const
  PropKeyMedia: array[TPropKeyMedia] of TPropertyItem = (
    (Key: @PKEY_Media_Duration;                  Name: 'Duration'),
    (Key: @PKEY_Media_ClassPrimaryID;            Name: 'ClassPrimaryID'),
    (Key: @PKEY_Media_ClassSecondaryID;          Name: 'ClassSecondaryID'),
    (Key: @PKEY_Media_DVDID;                     Name: 'DVDID'),
    (Key: @PKEY_Media_MCDI;                      Name: 'MCDI'),
    (Key: @PKEY_Media_MetadataContentProvider;   Name: 'MetadataContentProvider'),
    (Key: @PKEY_Media_ContentDistributor;        Name: 'ContentDistributor'),
    (Key: @PKEY_Media_Producer;                  Name: 'Producer'),
    (Key: @PKEY_Media_Writer;                    Name: 'Writer'),
    (Key: @PKEY_Media_CollectionGroupID;         Name: 'CollectionGroupID'),
    (Key: @PKEY_Media_CollectionID;              Name: 'CollectionID'),
    (Key: @PKEY_Media_ContentID;                 Name: 'ContentID'),
    (Key: @PKEY_Media_CreatorApplication;        Name: 'CreatorApplication'),
    (Key: @PKEY_Media_CreatorApplicationVersion; Name: 'CreatorApplicationVersion'),
    (Key: @PKEY_Media_Publisher;                 Name: 'Publisher'),
    (Key: @PKEY_Media_AuthorUrl;                 Name: 'AuthorUrl'),
    (Key: @PKEY_Media_PromotionUrl;              Name: 'PromotionUrl'),
    (Key: @PKEY_Media_UserWebUrl;                Name: 'UserWebUrl'),
    (Key: @PKEY_Media_UniqueFileIdentifier;      Name: 'UniqueFileIdentifier'),
    (Key: @PKEY_Media_EncodedBy;                 Name: 'EncodedBy'),
    (Key: @PKEY_Media_EncodingSettings;          Name: 'EncodingSettings'),
    (Key: @PKEY_Media_ProviderRating;            Name: 'ProviderRating'),
    (Key: @PKEY_Media_ProtectionType;            Name: 'ProtectionType'),
    (Key: @PKEY_Media_ProviderStyle;             Name: 'ProviderStyle'),
    (Key: @PKEY_Media_UserNoAutoInfo;            Name: 'UserNoAutoInfo'),

    (Key: @PKEY_Media_AverageLevel;              Name: 'AverageLevel'),
    (Key: @PKEY_Media_DateEncoded;               Name: 'DateEncoded'),
    (Key: @PKEY_Media_DateReleased;              Name: 'DateReleased'),
    (Key: @PKEY_Media_Year;                      Name: 'Year'),
    (Key: @PKEY_Media_SubTitle;                  Name: 'SubTitle'),
    (Key: @PKEY_Media_FrameCount;                Name: 'FrameCount'),
    (Key: @PKEY_Media_SubscriptionContentId;     Name: 'SubscriptionContentId')
  );


//
// PKEY_PropGroup_MediaAdvanced
//
type
  TPropKeyMediaAdvanced = (
    PK_MediaAdvanced_DRM_DatePlayExpires,
    PK_MediaAdvanced_DRM_DatePlayStarts,
    PK_MediaAdvanced_DRM_Description,
    PK_MediaAdvanced_DRM_IsProtected,
    PK_MediaAdvanced_DRM_PlayCount
  );
  TPropKeyMediaAdvancedFlags = set of TPropKeyMediaAdvanced;
const
  PropKeyMediaAdvanced: array[TPropKeyMediaAdvanced] of TPropertyItem = (
    (Key: @PKEY_DRM_DatePlayExpires; Name: 'DRM_DatePlayExpires'),
    (Key: @PKEY_DRM_DatePlayStarts;  Name: 'DRM_DatePlayStarts'),
    (Key: @PKEY_DRM_Description;     Name: 'DRM_Description'),
    (Key: @PKEY_DRM_IsProtected;     Name: 'DRM_IsProtected'),
    (Key: @PKEY_DRM_PlayCount;       Name: 'DRM_PlayCount')
  );


//
// PKEY_PropGroup_Message
//
type
  TPropKeyMessage = (
    PK_Message_AttachmentContents,
    PK_Message_AttachmentNames,
    PK_Message_BccAddress,
    PK_Message_BccName,
    PK_Message_CcAddress,
    PK_Message_CcName,
    PK_Message_ConversationID,
    PK_Message_ConversationIndex,
    PK_Message_DateReceived,
    PK_Message_DateSent,
    PK_Message_Flags,
    PK_Message_FromAddress,
    PK_Message_FromName,
    PK_Message_HasAttachments,
    PK_Message_IsFwdOrReply,
    PK_Message_MessageClass,
    PK_Message_ProofInProgress,
    PK_Message_SenderAddress,
    PK_Message_SenderName,
    PK_Message_Store,
    PK_Message_ToAddress,
    PK_Message_ToDoFlags,
    PK_Message_ToDoTitle,
    PK_Message_ToName
  );
  TPropKeyMessageFlags = set of TPropKeyMessage;
const
  PropKeyMessage: array[TPropKeyMessage] of TPropertyItem = (
    (Key: @PKEY_Message_AttachmentContents; Name: 'AttachmentContents'),
    (Key: @PKEY_Message_AttachmentNames;    Name: 'AttachmentNames'),
    (Key: @PKEY_Message_BccAddress;         Name: 'BccAddress'),
    (Key: @PKEY_Message_BccName;            Name: 'BccName'),
    (Key: @PKEY_Message_CcAddress;          Name: 'CcAddress'),
    (Key: @PKEY_Message_CcName;             Name: 'CcName'),
    (Key: @PKEY_Message_ConversationID;     Name: 'ConversationID'),
    (Key: @PKEY_Message_ConversationIndex;  Name: 'ConversationIndex'),
    (Key: @PKEY_Message_DateReceived;       Name: 'DateReceived'),
    (Key: @PKEY_Message_DateSent;           Name: 'DateSent'),
    (Key: @PKEY_Message_Flags;              Name: 'Flags'),
    (Key: @PKEY_Message_FromAddress;        Name: 'FromAddress'),
    (Key: @PKEY_Message_FromName;           Name: 'FromName'),
    (Key: @PKEY_Message_HasAttachments;     Name: 'HasAttachments'),
    (Key: @PKEY_Message_IsFwdOrReply;       Name: 'IsFwdOrReply'),
    (Key: @PKEY_Message_MessageClass;       Name: 'MessageClass'),
    (Key: @PKEY_Message_ProofInProgress;    Name: 'ProofInProgress'),
    (Key: @PKEY_Message_SenderAddress;      Name: 'SenderAddress'),
    (Key: @PKEY_Message_SenderName;         Name: 'SenderName'),
    (Key: @PKEY_Message_Store;              Name: 'Store'),
    (Key: @PKEY_Message_ToAddress;          Name: 'ToAddress'),
    (Key: @PKEY_Message_ToDoFlags;          Name: 'ToDoFlags'),
    (Key: @PKEY_Message_ToDoTitle;          Name: 'ToDoTitle'),
    (Key: @PKEY_Message_ToName;             Name: 'ToName')
  );


//
// PKEY_PropGroup_Music
//
type
  TPropKeyMusic = (
    PK_Music_AlbumArtist,
    PK_Music_AlbumID,
    PK_Music_AlbumTitle,
    PK_Music_Artist,
    PK_Music_BeatsPerMinute,
    PK_Music_Composer,
    PK_Music_Conductor,
    PK_Music_ContentGroupDescription,
    PK_Music_Genre,
    PK_Music_InitialKey,
    PK_Music_Lyrics,
    PK_Music_Mood,
    PK_Music_PartOfSet,
    PK_Music_Period,
    PK_Music_SynchronizedLyrics,
    PK_Music_TrackNumber
  );
  TPropKeyMusicFlags = set of TPropKeyMusic;
const
  PropKeyMusic: array[TPropKeyMusic] of TPropertyItem = (
    (Key: @PKEY_Music_AlbumArtist;             Name: 'AlbumArtist'),
    (Key: @PKEY_Music_AlbumID;                 Name: 'AlbumID'),
    (Key: @PKEY_Music_AlbumTitle;              Name: 'AlbumTitle'),
    (Key: @PKEY_Music_Artist;                  Name: 'Artist'),
    (Key: @PKEY_Music_BeatsPerMinute;          Name: 'BeatsPerMinute'),
    (Key: @PKEY_Music_Composer;                Name: 'Composer'),
    (Key: @PKEY_Music_Conductor;               Name: 'Conductor'),
    (Key: @PKEY_Music_ContentGroupDescription; Name: 'ContentGroupDescription'),
    (Key: @PKEY_Music_Genre;                   Name: 'Genre'),
    (Key: @PKEY_Music_InitialKey;              Name: 'InitialKey'),
    (Key: @PKEY_Music_Lyrics;                  Name: 'Lyrics'),
    (Key: @PKEY_Music_Mood;                    Name: 'Mood'),
    (Key: @PKEY_Music_PartOfSet;               Name: 'PartOfSet'),
    (Key: @PKEY_Music_Period;                  Name: 'Period'),
    (Key: @PKEY_Music_SynchronizedLyrics;      Name: 'SynchronizedLyrics'),
    (Key: @PKEY_Music_TrackNumber;             Name: 'TrackNumber')
  );


//
// PKEY_PropGroup_Origin
//
type
  TPropKeyOrigin = (
    PK_Document_ByteCount,
    PK_Document_CharacterCount,
    PK_Document_ClientID,
    PK_Document_Contributor,
    PK_Document_DateCreated,
    PK_Document_DatePrinted,
    PK_Document_DateSaved,
    PK_Document_Division,
    PK_Document_DocumentID,
    PK_Document_HiddenSlideCount,
    PK_Document_LastAuthor,
    PK_Document_LineCount,
    PK_Document_Manager,
    PK_Document_MultimediaClipCount,
    PK_Document_NoteCount,
    PK_Document_PageCount,
    PK_Document_ParagraphCount,
    PK_Document_PresentationFormat,
    PK_Document_RevisionNumber,
    PK_Document_Security,
    PK_Document_SlideCount,
    PK_Document_Template,
    PK_Document_TotalEditingTime,
    PK_Document_Version,
    PK_Document_WordCount
  );
  TPropKeyOriginFlags = set of TPropKeyOrigin;
const
  PropKeyOrigin: array[TPropKeyOrigin] of TPropertyItem = (
    (Key: @PKEY_Document_ByteCount;           Name: 'ByteCount'),
    (Key: @PKEY_Document_CharacterCount;      Name: 'CharacterCount'),
    (Key: @PKEY_Document_ClientID;            Name: 'ClientID'),
    (Key: @PKEY_Document_Contributor;         Name: 'Contributor'),
    (Key: @PKEY_Document_DateCreated;         Name: 'DateCreated'),
    (Key: @PKEY_Document_DatePrinted;         Name: 'DatePrinted'),
    (Key: @PKEY_Document_DateSaved;           Name: 'DateSaved'),
    (Key: @PKEY_Document_Division;            Name: 'Division'),
    (Key: @PKEY_Document_DocumentID;          Name: 'DocumentID'),
    (Key: @PKEY_Document_HiddenSlideCount;    Name: 'HiddenSlideCount'),
    (Key: @PKEY_Document_LastAuthor;          Name: 'LastAuthor'),
    (Key: @PKEY_Document_LineCount;           Name: 'LineCount'),
    (Key: @PKEY_Document_Manager;             Name: 'Manager'),
    (Key: @PKEY_Document_MultimediaClipCount; Name: 'MultimediaClipCount'),
    (Key: @PKEY_Document_NoteCount;           Name: 'NoteCount'),
    (Key: @PKEY_Document_PageCount;           Name: 'PageCount'),
    (Key: @PKEY_Document_ParagraphCount;      Name: 'ParagraphCount'),
    (Key: @PKEY_Document_PresentationFormat;  Name: 'PresentationFormat'),
    (Key: @PKEY_Document_RevisionNumber;      Name: 'RevisionNumber'),
    (Key: @PKEY_Document_Security;            Name: 'Security'),
    (Key: @PKEY_Document_SlideCount;          Name: 'SlideCount'),
    (Key: @PKEY_Document_Template;            Name: 'Template'),
    (Key: @PKEY_Document_TotalEditingTime;    Name: 'TotalEditingTime'),
    (Key: @PKEY_Document_Version;             Name: 'Version'),
    (Key: @PKEY_Document_WordCount;           Name: 'WordCount')
  );


//
// PKEY_PropGroup_PhotoAdvanced
//
type
  TPropKeyPhoto = (
    PK_Photo_Brightness,
    PK_Photo_BrightnessDenominator,
    PK_Photo_BrightnessNumerator,
    PK_Photo_Contrast,
    PK_Photo_ContrastText,
    PK_Photo_EXIFVersion,
    PK_Photo_GainControl,
    PK_Photo_GainControlDenominator,
    PK_Photo_GainControlNumerator,
    PK_Photo_GainControlText,
    PK_Photo_MakerNote,
    PK_Photo_MakerNoteOffset,
    PK_Photo_PeopleNames,
    PK_Photo_PhotometricInterpretation,
    PK_Photo_PhotometricInterpretationText,
    PK_Photo_ProgramMode,
    PK_Photo_ProgramModeText,
    PK_Photo_RelatedSoundFile,
    PK_Photo_Saturation,
    PK_Photo_SaturationText,
    PK_Photo_Sharpness,
    PK_Photo_SharpnessText,
    PK_Photo_TagViewAggregate,
    PK_Photo_TranscodedForSync,
    PK_Photo_WhiteBalance,
    PK_Photo_WhiteBalanceText
  );
  TPropKeyPhotoFlags = set of TPropKeyPhoto;
const
  PropKeyPhoto: array[TPropKeyPhoto] of TPropertyItem = (
    (Key: @PKEY_Photo_Brightness;             Name: 'Brightness'),
    (Key: @PKEY_Photo_BrightnessDenominator;  Name: 'BrightnessDenominator'),
    (Key: @PKEY_Photo_BrightnessNumerator;    Name: 'BrightnessNumerator'),
    (Key: @PKEY_Photo_Contrast;               Name: 'Contrast'),
    (Key: @PKEY_Photo_ContrastText;           Name: 'ContrastText'),
    (Key: @PKEY_Photo_EXIFVersion;            Name: 'EXIFVersion'),
    (Key: @PKEY_Photo_GainControl;            Name: 'GainControl'),
    (Key: @PKEY_Photo_GainControlDenominator; Name: 'GainControlDenominator'),
    (Key: @PKEY_Photo_GainControlNumerator;   Name: 'GainControlNumerator'),
    (Key: @PKEY_Photo_GainControlText;        Name: 'GainControlText'),
    (Key: @PKEY_Photo_MakerNote;              Name: 'MakerNote'),
    (Key: @PKEY_Photo_MakerNoteOffset;        Name: 'MakerNoteOffset'),
    (Key: @PKEY_Photo_PeopleNames;            Name: 'PeopleNames'),
    (Key: @PKEY_Photo_PhotometricInterpretation;     Name: 'PhotometricInterpretation'),
    (Key: @PKEY_Photo_PhotometricInterpretationText; Name: 'PhotometricInterpretationText'),
    (Key: @PKEY_Photo_ProgramMode;            Name: 'ProgramMode'),
    (Key: @PKEY_Photo_ProgramModeText;        Name: 'ProgramModeText'),
    (Key: @PKEY_Photo_RelatedSoundFile;       Name: 'RelatedSoundFile'),
    (Key: @PKEY_Photo_Saturation;             Name: 'Saturation'),
    (Key: @PKEY_Photo_SaturationText;         Name: 'SaturationText'),
    (Key: @PKEY_Photo_Sharpness;              Name: 'Sharpness'),
    (Key: @PKEY_Photo_SharpnessText;          Name: 'SharpnessText'),
    (Key: @PKEY_Photo_TagViewAggregate;       Name: 'TagViewAggregate'),
    (Key: @PKEY_Photo_TranscodedForSync;      Name: 'TranscodedForSync'),
    (Key: @PKEY_Photo_WhiteBalance;           Name: 'WhiteBalance'),
    (Key: @PKEY_Photo_WhiteBalanceText;       Name: 'WhiteBalanceText')
  );


//
// PKEY_PropGroup_RecordedTV
//
type
  TPropKeyRecordedTV = (
    PK_RecordedTV_ChannelNumber,
    PK_RecordedTV_Credits,
    PK_RecordedTV_DateContentExpires,
    PK_RecordedTV_EpisodeName,
    PK_RecordedTV_IsATSCContent,
    PK_RecordedTV_IsClosedCaptioningAvailable,
    PK_RecordedTV_IsDTVContent,
    PK_RecordedTV_IsHDContent,
    PK_RecordedTV_IsRepeatBroadcast,
    PK_RecordedTV_IsSAP,
    PK_RecordedTV_NetworkAffiliation,
    PK_RecordedTV_OriginalBroadcastDate,
    PK_RecordedTV_ProgramDescription,
    PK_RecordedTV_RecordingTime,
    PK_RecordedTV_StationCallSign,
    PK_RecordedTV_StationName
  );
  TPropKeyRecordedTVFlags = set of TPropKeyRecordedTV;
const
  PropKeyRecordedTV: array[TPropKeyRecordedTV] of TPropertyItem = (
    (Key: @PKEY_RecordedTV_ChannelNumber;         Name: 'ChannelNumber'),
    (Key: @PKEY_RecordedTV_Credits;               Name: 'Credits'),
    (Key: @PKEY_RecordedTV_DateContentExpires;    Name: 'DateContentExpires'),
    (Key: @PKEY_RecordedTV_EpisodeName;           Name: 'EpisodeName'),
    (Key: @PKEY_RecordedTV_IsATSCContent;         Name: 'IsATSCContent'),
    (Key: @PKEY_RecordedTV_IsClosedCaptioningAvailable; Name: 'IsClosedCaptioningAvailable'),
    (Key: @PKEY_RecordedTV_IsDTVContent;          Name: 'IsDTVContent'),
    (Key: @PKEY_RecordedTV_IsHDContent;           Name: 'IsHDContent'),
    (Key: @PKEY_RecordedTV_IsRepeatBroadcast;     Name: 'IsRepeatBroadcast'),
    (Key: @PKEY_RecordedTV_IsSAP;                 Name: 'IsSAP'),
    (Key: @PKEY_RecordedTV_NetworkAffiliation;    Name: 'NetworkAffiliation'),
    (Key: @PKEY_RecordedTV_OriginalBroadcastDate; Name: 'OriginalBroadcastDate'),
    (Key: @PKEY_RecordedTV_ProgramDescription;    Name: 'ProgramDescription'),
    (Key: @PKEY_RecordedTV_RecordingTime;         Name: 'RecordingTime'),
    (Key: @PKEY_RecordedTV_StationCallSign;       Name: 'StationCallSign'),
    (Key: @PKEY_RecordedTV_StationName;           Name: 'StationName')
  );


//
// PKEY_PropGroup_Video
//
type
  TPropKeyVideo = (
    PK_Video_StreamName,
    PK_Video_FrameWidth,
    PK_Video_FrameHeight,
    PK_Video_FrameRate,
    PK_Video_EncodingBitrate,
    PK_Video_SampleSize,
    PK_Video_Compression,
    PK_Video_StreamNumber,
    PK_Video_Director,
    PK_Video_HorizontalAspectRatio,
    PK_Video_TotalBitrate,
    PK_Video_FourCC,
    PK_Video_VerticalAspectRatio
  );
  TPropKeyVideoFlags = set of TPropKeyVideo;
const
  PropKeyVideo: array[TPropKeyVideo] of TPropertyItem = (
    (Key: @PKEY_Video_StreamName;            Name: 'StreamName'),
    (Key: @PKEY_Video_FrameWidth;            Name: 'FrameWidth'),
    (Key: @PKEY_Video_FrameHeight;           Name: 'FrameHeight'),
    (Key: @PKEY_Video_FrameRate;             Name: 'FrameRate'),
    (Key: @PKEY_Video_EncodingBitrate;       Name: 'EncodingBitrate'),
    (Key: @PKEY_Video_SampleSize;            Name: 'SampleSize'),
    (Key: @PKEY_Video_Compression;           Name: 'Compression'),
    (Key: @PKEY_Video_StreamNumber;          Name: 'StreamNumber'),
    (Key: @PKEY_Video_Director;              Name: 'Director'),
    (Key: @PKEY_Video_HorizontalAspectRatio; Name: 'HorizontalAspectRatio'),
    (Key: @PKEY_Video_TotalBitrate;          Name: 'TotalBitrate'),
    (Key: @PKEY_Video_FourCC;                Name: 'FourCC'),
    (Key: @PKEY_Video_VerticalAspectRatio;   Name: 'VerticalAspectRatio')
  );

// PKEY_PropGroup_Advanced


//
// Group
// 檔案屬性的詳細資料群組，自訂屬性的分類，與系統的不完全相同
//
type
  TPropKeyGroup = (
    PK_PropGroup_General, // 檔案屬性的一般標籤頁
    PK_PropGroup_FileSystem,
    PK_PropGroup_Description,
    PK_PropGroup_Content,
    PK_PropGroup_Origin,
    PK_PropGroup_Calendar,
    PK_PropGroup_Media,
    PK_PropGroup_MediaAdvanced,
    PK_PropGroup_Image,
    PK_PropGroup_PhotoAdvanced,
    PK_PropGroup_RecordedTV,
    PK_PropGroup_Camera,
    PK_PropGroup_Video,
    PK_PropGroup_Music,
    PK_PropGroup_Audio,
    PK_PropGroup_GPS,
    PK_PropGroup_Contact,
    PK_PropGroup_Message,
    PK_PropGroup_Advanced
  );
  TPropertyGroupItem = record
    Key: PPropertyKey;
    Count: Integer;
    Items: PPropertyItem;
    Name: string;
  end;
  PPropertyGroupItem = ^TPropertyGroupItem;
const
  PropKeyGroup: array[TPropKeyGroup] of TPropertyGroupItem = (
    (Key: @PKEY_PropGroup_General;      Count: Length(PropKeyGeneral);      Items: @PropKeyGeneral;      Name: 'General'),
    (Key: @PKEY_PropGroup_FileSystem;   Count: Length(PropKeyFileSystem);   Items: @PropKeyFileSystem;   Name: 'FileSystem'),
    (Key: @PKEY_PropGroup_Description;  Count: Length(PropKeyDescription);  Items: @PropKeyDescription;  Name: 'Description'),
    (Key: @PKEY_PropGroup_Content;      Count: Length(PropKeyContent);      Items: @PropKeyContent;      Name: 'Content'),
    (Key: @PKEY_PropGroup_Origin;       Count: Length(PropKeyOrigin);       Items: @PropKeyOrigin;       Name: 'Origin'),
    (Key: @PKEY_PropGroup_Calendar;     Count: Length(PropKeyCalendar);     Items: @PropKeyCalendar;     Name: 'Calendar'),
    (Key: @PKEY_PropGroup_Media;        Count: Length(PropKeyMedia);        Items: @PropKeyMedia;        Name: 'Media'),
    (Key: @PKEY_PropGroup_MediaAdvanced;Count: Length(PropKeyMediaAdvanced);Items: @PropKeyMediaAdvanced;Name: 'MediaAdvanced'),
    (Key: @PKEY_PropGroup_Image;        Count: Length(PropKeyImage);        Items: @PropKeyImage;        Name: 'Image'),
    (Key: @PKEY_PropGroup_PhotoAdvanced;Count: Length(PropKeyPhoto);        Items: @PropKeyPhoto;        Name: 'PhotoAdvanced'),
    (Key: @PKEY_PropGroup_RecordedTV;   Count: Length(PropKeyRecordedTV);   Items: @PropKeyRecordedTV;   Name: 'RecordedTV'),
    (Key: @PKEY_PropGroup_Camera;       Count: Length(PropKeyCamera);       Items: @PropKeyCamera;       Name: 'Camera'),
    (Key: @PKEY_PropGroup_Video;        Count: Length(PropKeyVideo);        Items: @PropKeyVideo;        Name: 'Video'),
    (Key: @PKEY_PropGroup_Music;        Count: Length(PropKeyMusic);        Items: @PropKeyMusic;        Name: 'Music'),
    (Key: @PKEY_PropGroup_Audio;        Count: Length(PropKeyVideo);        Items: @PropKeyAudio;        Name: 'Audio'),
    (Key: @PKEY_PropGroup_GPS;          Count: Length(PropKeyGPS);          Items: @PropKeyGPS;          Name: 'GPS'),
    (Key: @PKEY_PropGroup_Contact;      Count: Length(PropKeyContact);      Items: @PropKeyContact;      Name: 'Contact'),
    (Key: @PKEY_PropGroup_Message;      Count: Length(PropKeyMessage);      Items: @PropKeyMessage;      Name: 'Message'),
    (Key: @PKEY_PropGroup_Advanced;     Count: 0;                           Items: nil;                  Name: 'Advanced')
  );


//
// Error return codes
//
const
  STRSAFE_E_INSUFFICIENT_BUFFER = HResult($8007007A);
  STRSAFE_E_INVALID_PARAMETER   = HResult($80070057);
  STRSAFE_E_END_OF_FILE         = HResult($80070026);


//
// 功能函數
//

// 搜尋目標項目的索引
function IndexOfPropItem(Items: PPropertyItem; Count: Integer; const Key: TPropertyKey): Integer;
function IndexOfGroup(const Key: TPropertyKey; var iItem: Integer): Integer;


//
// 簡化取得系統屬性資訊
//

// 取得 介面 的屬性描述
function GetDescription(const propkey: TPropertyKey; var PropDesc: IPropertyDescription): Boolean;

// 取得 介面 的顯示名稱，依照系統地區語言
function GetDisplayName(const propkey: TPropertyKey; var Text: string): Boolean; overload;
function GetDisplayName(const propkey: TPropertyKey): string; overload;

// 取得 正式名稱 (正準名稱)
function GetCanonicalName(const propkey: TPropertyKey; var Text: string): Boolean; overload;
function GetCanonicalName(const propkey: TPropertyKey): string; overload;

implementation

//
// 搜尋目標項目的索引
//
function IndexOfPropItem(Items: PPropertyItem; Count: Integer; const Key: TPropertyKey): Integer;
var
  I: Integer;
begin
  I := 0;
  while I < Count do
  begin
    if Items.Key^ = Key then
      Exit(I);
    Inc(Items);
    Inc(I);
  end;
  Result := -1;
end;

function IndexOfGroup(const Key: TPropertyKey; var iItem: Integer): Integer;
var
  I, J, Count: Integer;
  Group: PPropertyGroupItem;
begin
  I := 0;
  Group := @PropKeyGroup;
  Count := Length(PropKeyGroup);
  while I < Count do
  begin
    J := IndexOfPropItem(Group.Items, Group.Count, Key);
    if J >= 0 then
    begin
      iItem := J;
      Exit(I);
    end;
    Inc(Group);
    Inc(I);
  end;
  Result := -1;
end;


//
// 簡化取得系統屬性資訊
//

// 取得 介面 的屬性描述
function GetDescription(const propkey: TPropertyKey; var PropDesc: IPropertyDescription): Boolean;
var
  hR: HResult;
begin
  hR := PSGetPropertyDescription(propkey, GetTypeData(TypeInfo(IPropertyDescription))^.Guid, Pointer(PropDesc));
  Result := Succeeded(hR);
end;

// 取得 正式名稱 (正準名稱)
function GetCanonicalName(const propkey: TPropertyKey; var Text: string): Boolean;
var
  hR: HResult;
  pStr: PWideChar;
  PropDesc:IPropertyDescription;
begin
  pStr := nil;
  if GetDescription(propkey, PropDesc) then
    try
      hR := PropDesc.GetCanonicalName(pStr);
      if Succeeded(hR) then
        try
          Text := WideCharToString(pStr);
          Exit(True);
        finally
          CoTaskMemFree(pStr); // PSGetNameFromPropertyKey
        end;
    finally
      PropDesc := nil;
    end;
  Result := False;
end;

// 取得 正式名稱 (正準名稱)
function GetCanonicalName(const propkey: TPropertyKey): string;
begin
  if not GetCanonicalName(propkey, Result) then
    Result := '';
end;

// 取得顯示的名稱，依照系統地區語言
function GetDisplayName(const propkey: TPropertyKey; var Text: string): Boolean;
var
  hR: HResult;
  pStr: PWideChar;
  PropDesc:IPropertyDescription;
begin
  pStr := nil;
  if GetDescription(propkey, PropDesc) then
    try
      hR := PropDesc.GetDisplayName(pStr);
      if Succeeded(hR) then
        try
          Text := WideCharToString(pStr);
          Exit(True);
        finally
          CoTaskMemFree(pStr); // PSGetNameFromPropertyKey
        end;
    finally
      PropDesc := nil;
    end;
  Result := False;
end;

// 取得顯示的名稱，依照系統地區語言
function GetDisplayName(const propkey: TPropertyKey): string;
begin
  if not GetDisplayName(propkey, Result) then
    Result := '';
end;

{ tagPROPVARIANTHelper }

// 利用系統內建函數轉換 PropVariant 內容為字串型別
function tagPROPVARIANTHelper.TryToString(var str: string): Boolean;
var
  hR: HResult;
  Value: TPropVariant;
begin
  hR := PropVariantChangeType(Value, Self, 0, VT_LPWSTR);
  Result := Succeeded(hR);
  if Result then
  begin
    str := WideCharToString(Value.bstrVal);
    PropVariantClear(Value);
  end;
end;

// 取得轉換 PropVariant 內容後的字串訊息
function tagPROPVARIANTHelper.ToString: string;
begin
  if not TryToString(Result) then
    Result := '';
end;

end.
