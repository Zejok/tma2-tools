"^(?:See )?(wbem\w*|WBEM\w*)(?:\s.*)?\r^(?:\d{1,10})\s\(0x([A-Z0-9]{1,8})\)\s*^\s*(.+)*$"

;The call was successful.
Const $wbemNoErr = 0x0
;The call failed.
Const $wbemErrFailed = 0x80041001
;The object could not be found.
Const $wbemErrNotFound = 0x80041002
;The current user does not have permission to perform the action.
Const $wbemErrAccessDenied = 0x80041003
;The provider has failed at some time other than during initialization.
Const $wbemErrProviderFailure = 0x80041004
;A type mismatch occurred.
Const $wbemErrTypeMismatch = 0x80041005
;There was not enough memory for the operation.
Const $wbemErrOutOfMemory = 0x80041006
;The SWbemNamedValue object is not valid.
Const $wbemErrInvalidContext = 0x80041007
;One of the parameters to the call is not correct.
Const $wbemErrInvalidParameter = 0x80041008
;The resource, typically a remote server, is not currently available.
Const $wbemErrNotAvailable = 0x80041009
;An internal, critical, and unexpected error occurred. Report this error to Microsoft Technical Support.
Const $wbemErrCriticalError = 0x8004100A
;One or more network packets were corrupted during a remote session.
Const $wbemErrInvalidStream = 0x8004100B
;The feature or operation is not supported.
Const $wbemErrNotSupported = 0x8004100C
;The parent class specified is not valid.
Const $wbemErrInvalidSuperclass = 0x8004100D
;The namespace specified could not be found.
Const $wbemErrInvalidNamespace = 0x8004100E
;The specified instance is not valid.
Const $wbemErrInvalidObject = 0x8004100F
;The specified class is not valid.
Const $wbemErrInvalidClass = 0x80041010
;A provider referenced in the schema does not have a corresponding registration.
Const $wbemErrProviderNotFound = 0x80041011
;A provider referenced in the schema has an incorrect or incomplete registration. This error may be caused by a missing pragma namespace command in the MOF file used to register the provider, resulting in the provider being registered in the wrong WMI namespace. This error may also be caused by a corrupt repository, which may be fixed by deleting it and recompiling the MOF files.
Const $wbemErrInvalidProviderRegistration = 0x80041012
;COM cannot locate a provider referenced in the schema. This error may be caused by any of the following:
Const $wbemErrProviderLoadFailure = 0x80041013
The provider is using a WMI DLL that does not match the .lib fileused when the provider was built.

The provider's DLL or any of the DLLs on which it depends is corrupt.

The provider failed to export DllRegisterServer.

An in-process provider was not registered using /regsvr32.

An out-of-process provider was not registered using /regserver.

;A component, such as a provider, failed to initialize for internal reasons.
Const $wbemErrInitializationFailure = 0x80041014
;A networking error occurred, preventing normal operation.
Const $wbemErrTransportFailure = 0x80041015
;The requested operation is not valid. This error usually applies to invalid attempts to delete classes or properties.
Const $wbemErrInvalidOperation = 0x80041016
;The requested operation is not valid. This error usually applies to invalid attempts to delete classes or properties.
Const $wbemErrInvalidQuery = 0x80041017
;The requested query language is not supported.
Const $wbemErrInvalidQueryType = 0x80041018
;In a put operation, the wbemChangeFlagCreateOnly flag was specified, but the instance already exists.
Const $wbemErrAlreadyExists = 0x80041019
;It is not possible to perform the add operation on this qualifier because the owning object does not permit overrides.
Const $wbemErrOverrideNotAllowed = 0x8004101A
;The user attempted to delete a qualifier that was not owned. The qualifier was inherited from a parent class.
Const $wbemErrPropagatedQualifier = 0x8004101B
;The user attempted to delete a property that was not owned. The property was inherited from a parent class.
Const $wbemErrPropagatedProperty = 0x8004101C
;The client made an unexpected and illegal sequence of calls, such as calling EndEnumeration before calling BeginEnumeration.
Const $wbemErrUnexpected = 0x8004101D
;The user requested an illegal operation, such as spawning a class from an instance.
Const $wbemErrIllegalOperation = 0x8004101E
;There was an illegal attempt to specify a key qualifier on a property that cannot be a key. The keys are specified in the class definition for an object, and cannot be altered on a per-instance basis.
Const $wbemErrCannotBeKey = 0x8004101F
;The current object is not a valid class definition. Either it is incomplete, or it has not been registered with WMI using SWbemObject.Put_.
Const $wbemErrIncompleteClass = 0x80041020
;The syntax of an input parameter is incorrect for the applicable data structure. For example, when a CIM datetime structure does not have the correct format when passed to SWbemDateTime.SetFileTime.
Const $wbemErrInvalidSyntax = 0x80041021
;Reserved for future use.
Const $wbemErrNondecoratedObject = 0x80041022
;The property that you are attempting to modify is read-only.
Const $wbemErrReadOnly = 0x80041023
;The provider cannot perform the requested operation. This would include a query that is too complex, retrieving an instance, creating or updating a class, deleting a class, or enumerating a class.
Const $wbemErrProviderNotCapable = 0x80041024
;An attempt was made to make a change that would invalidate a subclass.
Const $wbemErrClassHasChildren = 0x80041025
;An attempt has been made to delete or modify a class that has instances.
Const $wbemErrClassHasInstances = 0x80041026
;Reserved for future use.
Const $wbemErrQueryNotImplemented = 0x80041027
;A value of Nothing was specified for a property that may not be Nothing, such as one that is marked by a Key, Indexed, or Not_Null qualifier.
Const $wbemErrIllegalNull = 0x80041028
;The CIM type specified for a property is not valid.
Const $wbemErrInvalidQualifierType = 0x80041029
;The CIM type specified for a property is not valid.
Const $wbemErrInvalidPropertyType = 0x8004102A
;The request was made with an out-of-range value, or is incompatible with the type.
Const $wbemErrValueOutOfRange = 0x8004102B
;An illegal attempt was made to make a class singleton, such as when the class is derived from a non-singleton class.
Const $wbemErrCannotBeSingleton = 0x8004102C
;The CIM type specified is not valid.
Const $wbemErrInvalidCimType = 0x8004102D
;The requested method is not available.
Const $wbemErrInvalidMethod = 0x8004102E
;The parameters provided for the method are not valid.
Const $wbemErrInvalidMethodParameters = 0x8004102F
;There was an attempt to get qualifiers on a system property.
Const $wbemErrSystemProperty = 0x80041030
;The property type is not recognized.
Const $wbemErrInvalidProperty = 0x80041031
;An asynchronous process has been canceled internally or by the user. Note that due to the timing and nature of the asynchronous operation the operation may not have been truly canceled.
Const $wbemErrCallCancelled = 0x80041032
;The user has requested an operation while WMI is in the process of shutting down.
Const $wbemErrShuttingDown = 0x80041033
;An attempt was made to reuse an existing method name from a parent class, and the signatures did not match.
Const $wbemErrPropagatedMethod = 0x80041034
;One or more parameter values, such as a query text, is too complex or unsupported. WMI is therefore requested to retry the operation with simpler parameters.
Const $wbemErrUnsupportedParameter = 0x80041035
;A parameter was missing from the method call.
Const $wbemErrMissingParameter = 0x80041036
;A method parameter has an invalid ID qualifier.
Const $wbemErrInvalidParameterId = 0x80041037
;One or more of the method parameters have ID qualifiers that are out of sequence.
Const $wbemErrNonConsecutiveParameterIds = 0x80041038
;The return value for a method has an ID qualifier.
Const $wbemErrParameterIdOnRetval = 0x80041039
;The specified object path was invalid.
Const $wbemErrInvalidObjectPath = 0x8004103A
wbemErrOutOfDiskSpace
2147749947



    Windows XP/2000/NT:  Disk is out of space.

    Windows Server 2003:  Disk is out of space or the 4 GB limit on WMI repository (CIM repository) size is reached.

;The supplied buffer was too small to hold all the objects in the enumerator or to read a string property.
Const $wbemErrBufferTooSmall = 0x8004103C
;The provider does not support the requested put operation.
Const $wbemErrUnsupportedPutExtension = 0x8004103D
;An object with an incorrect type or version was encountered during marshaling.
Const $wbemErrUnknownObjectType = 0x8004103E
;A packet with an incorrect type or version was encountered during marshaling.
Const $wbemErrUnknownPacketType = 0x8004103F
;The packet has an unsupported version.
Const $wbemErrMarshalVersionMismatch = 0x80041040
;The packet appears to be corrupted.
Const $wbemErrMarshalInvalidSignature = 0x80041041
;An attempt has been made to mismatch qualifiers, such as putting [key] on an object instead of a property.
Const $wbemErrInvalidQualifier = 0x80041042
;A duplicate parameter has been declared in a CIM method.
Const $wbemErrInvalidDuplicateParameter = 0x80041043
;Reserved for future use.
Const $wbemErrTooMuchData = 0x80041044
;A call to IWbemObjectSink::Indicate has failed. The provider may choose to refire the event.
Const $wbemErrServerTooBusy = 0x80041045
;The specified flavor was invalid.
Const $wbemErrInvalidFlavor = 0x80041046
;An attempt has been made to create a reference that is circular (for example, deriving a class from itself).
Const $wbemErrCircularReference = 0x80041047
wbemErrUnsupportedClassUpdate
2147749960



The specified class is not supported.

;An attempt was made to change a key when instances or subclasses are already using the key.
Const $wbemErrCannotChangeKeyInheritance = 0x80041049
;An attempt was made to change an index when instances or subclasses are already using the index.
Const $wbemErrCannotChangeIndexInheritance = 0x80041050
;An attempt was made to create more properties than the current version of the class supports.
Const $wbemErrTooManyProperties = 0x80041051
;A property was redefined with a conflicting type in a derived class.
Const $wbemErrUpdateTypeMismatch = 0x80041052
;An attempt was made in a derived class to override a non-overrideable qualifier.
Const $wbemErrUpdateOverrideNotAllowed = 0x80041053
;A method was redeclared with a conflicting signature in a derived class.
Const $wbemErrUpdatePropagatedMethod = 0x80041054
;An attempt was made to execute a method not marked with [implemented] in any relevant class.
Const $wbemErrMethodNotImplemented = 0x80041055
;An attempt was made to execute a method marked with [disabled].
Const $wbemErrMethodDisabled = 0x80041056
;The refresher is busy with another operation.
Const $wbemErrRefresherBusy = 0x80041057
;The filtering query is syntactically invalid.
Const $wbemErrUnparsableQuery = 0x80041058
;The FROM clause of a filtering query references a class that is not an event class (not derived from __Event).
Const $wbemErrNotEventClass = 0x80041059
;A GROUP BY clause was used without the corresponding GROUP WITHIN clause.
Const $wbemErrMissingGroupWithin = 0x8004105A
;A GROUP BY clause was used. Aggregation on all properties is not supported.
Const $wbemErrMissingAggregationList = 0x8004105B
;Dot notation was used on a property that is not an embedded object.
Const $wbemErrPropertyNotAnObject = 0x8004105C
;A GROUP BY clause references a property that is an embedded object without using dot notation.
Const $wbemErrAggregatingByObject = 0x8004105D
;An event provider registration query ( __EventProviderRegistration) did not specify the classes for which events were provided.
Const $wbemErrUninterpretableProviderQuery = 0x8004105F
;An request was made to back up or restore the repository while WMI was using it.
Const $wbemErrBackupRestoreWinmgmtRunning = 0x80041060
;The asynchronous delivery queue overflowed due to the event consumer being too slow.
Const $wbemErrQueueOverflow = 0x80041061
;The operation failed because the client did not have the necessary security privilege.
Const $wbemErrPrivilegeNotHeld = 0x80041062
;The operator is not valid for this property type.
Const $wbemErrInvalidOperator = 0x80041063
;The user specified a username, password or authority for a local connection. The user must use a blank username/password and rely on default security.
Const $wbemErrLocalCredentials = 0x80041064
;The class was made abstract when its parent class is not abstract.
Const $wbemErrCannotBeAbstract = 0x80041065
;An amended object was put without the wbemFlagUseAmendedQualifiers flag being specified.
Const $wbemErrAmendedObject = 0x80041066
;    Windows Server 2003 and Windows XP:  The client was not retrieving objects quickly enough from an enumeration. This constant is returned when a client creates an enumeration object but does not retrieve objects from the enumerator in a timely fashion, causing the enumerator's object caches to get backed up.
Const $wbemErrClientTooSlow = 0x80041067
;    Windows Server 2003 and Windows XP:  A null security descriptor was used.
Const $wbemErrNullSecurityDescriptor = 0x80041068
;    Windows Server 2003 and Windows XP:  The operation timed out.
Const $wbemErrTimeout = 0x80041069
;    Windows Server 2003 and Windows XP:  The association being used is not valid.
Const $wbemErrInvalidAssociation = 0x8004106A
;    Windows Server 2003 and Windows XP:  The operation was ambiguous.
Const $wbemErrAmbiguousOperation = 0x8004106B
;    Windows Server 2003 and Windows XP:  WMI is taking up too much memory. This could be caused either by low memory availability or excessive memory consumption by WMI.
Const $wbemErrQuotaViolation = 0x8004106C
;    Windows Server 2003 and Windows XP:  The operation resulted in a transaction conflict.
Const $wbemErrTransactionConflict = 0x8004106D
;    Windows Server 2003 and Windows XP:  The transaction forced a rollback.
Const $wbemErrForcedRollback = 0x8004106E
;    Windows Server 2003 and Windows XP:  The locale used in the call is not supported.
Const $wbemErrUnsupportedLocale = 0x8004106F
;    Windows Server 2003 and Windows XP:  The object handle is out of date.
Const $wbemErrHandleOutOfDate = 0x80041070
;    Windows Server 2003 and Windows XP:  Indicates that the connection to the SQL database failed.
Const $wbemErrConnectionFailed = 0x80041071
;    Windows Server 2003 and Windows XP:  The handle request was invalid.
Const $wbemErrInvalidHandleRequest = 0x80041072
;    Windows Server 2003 and Windows XP:  The property name contains more than 255 characters.
Const $wbemErrPropertyNameTooWide = 0x80041073
;    Windows Server 2003 and Windows XP:  The class name contains more than 255 characters.
Const $wbemErrClassNameTooWide = 0x80041074
;    Windows Server 2003 and Windows XP:  The method name contains more than 255 characters.
Const $wbemErrMethodNameTooWide = 0x80041075
;    Windows Server 2003 and Windows XP:  The qualifier name contains more than 255 characters.
Const $wbemErrQualifierNameTooWide = 0x80041076
;    Windows Server 2003 and Windows XP:  Indicates that an SQL command should be rerun because there is a deadlock in SQL. This can be returned only when data is being stored in an SQL database.
Const $wbemErrRerunCommand = 0x80041077
;    Windows Server 2003 and Windows XP:  The database version does not match the version that the repository driver understands.
Const $wbemErrDatabaseVerMismatch = 0x80041078
;    Windows Server 2003 and Windows XP:  WMI cannot do the delete operation because the provider does not allow it.
Const $wbemErrVetoDelete = 0x8004107A
;    Windows Server 2003 and Windows XP:  WMI cannot do the put operation because the provider does not allow it.
Const $wbemErrVetoPut = 0x8004107A
;    Windows Server 2003 and Windows XP:  The specified locale identifier was not valid for the operation.
Const $wbemErrInvalidLocale = 0x80041080
;    Windows Server 2003 and Windows XP:  The provider is suspended.
Const $wbemErrProviderSuspended = 0x80041081
;    Windows Server 2003 and Windows XP:  The object must be committed and retrieved again before the requested operation can succeed. This constant is returned when an object must be committed and re-retrieved to see the property value.
Const $wbemErrSynchronizationRequired = 0x80041082
;    Windows Server 2003 and Windows XP:  The operation cannot be completed because no schema is available.
Const $wbemErrNoSchema = 0x80041083
;    Windows Server 2003 and Windows XP:  The provider registration cannot be done because the provider is already registered.
Const $wbemErrProviderAlreadyRegistered = 0x80041084
;    Windows Server 2003 and Windows XP:  The provider for the requested data is not registered.
Const $wbemErrProviderNotRegistered = 0x80041085
;    Windows Server 2003 and Windows XP:  A fatal transport error occurred and other transport will not be attempted.
Const $wbemErrFatalTransportError = 0x80041086
;    Windows Server 2003 and Windows XP:  The client connection to WINMGMT must be encrypted for this operation. The IWbemServices proxy security settings should be adjusted and the operation retried.
Const $wbemErrEncryptedConnectionRequired = 0x80041087
;    Windows Server 2003 and Windows XP:  A provider failed to report results within the specified timeout.
Const $WBEM_E_PROVIDER_TIMED_OUT = 0x80041088
;    Windows Server 2003 and Windows XP:  User attempted to put an instance with no defined key.
Const $WBEM_E_NO_KEY = 0x80041089
;    Windows Server 2003 and Windows XP:  User attempted to register a provider instance but the COM server for the provider instance was unloaded.
Const $WBEM_E_PROVIDER_DISABLED = 0x8004108A
;    Windows Server 2003 and Windows XP:  The provider registration overlaps with the system event domain.
Const $wbemErrRegistrationTooBroad = 0x80042001
;    Windows Server 2003 and Windows XP:  A WITHIN clause was not used in this query.
Const $wbemErrRegistrationTooPrecise = 0x80042002
;    Windows Server 2003 and Windows XP:  Automation-specific error.
Const $wbemErrTimedout = 0x80043001
wbemErrResetToDefault
2147758082 (0x80043002)



    Windows Server 2003 and Windows XP:  The user deleted an override default value for the current class. The default value for this property in the parent class has been reactivated. An automation-specific error.

#region dictionary sections
$WbemErrorEnum.Add("wbemNoErr", Dec(0x0))

$WbemErrorEnum.Add("wbemErrFailed", Dec(0x80041001))

$WbemErrorEnum.Add("wbemErrNotFound", Dec(0x80041002))

$WbemErrorEnum.Add("wbemErrAccessDenied", Dec(0x80041003))

$WbemErrorEnum.Add("wbemErrProviderFailure", Dec(0x80041004))

$WbemErrorEnum.Add("wbemErrTypeMismatch", Dec(0x80041005))

$WbemErrorEnum.Add("wbemErrOutOfMemory", Dec(0x80041006))

$WbemErrorEnum.Add("wbemErrInvalidContext", Dec(0x80041007))

$WbemErrorEnum.Add("wbemErrInvalidParameter", Dec(0x80041008))

$WbemErrorEnum.Add("wbemErrNotAvailable", Dec(0x80041009))

$WbemErrorEnum.Add("wbemErrCriticalError", Dec(0x8004100A))

$WbemErrorEnum.Add("wbemErrInvalidStream", Dec(0x8004100B))

$WbemErrorEnum.Add("wbemErrNotSupported", Dec(0x8004100C))

$WbemErrorEnum.Add("wbemErrInvalidSuperclass", Dec(0x8004100D))

$WbemErrorEnum.Add("wbemErrInvalidNamespace", Dec(0x8004100E))

$WbemErrorEnum.Add("wbemErrInvalidObject", Dec(0x8004100F))

$WbemErrorEnum.Add("wbemErrInvalidClass", Dec(0x80041010))

$WbemErrorEnum.Add("wbemErrProviderNotFound", Dec(0x80041011))

$WbemErrorEnum.Add("wbemErrInvalidProviderRegistration", Dec(0x80041012))

$WbemErrorEnum.Add("wbemErrProviderLoadFailure", Dec(0x80041013))

The provider is using a WMI DLL that does not match the .lib fileused when the provider was built.

The provider's DLL or any of the DLLs on which it depends is corrupt.

The provider failed to export DllRegisterServer.

An in-process provider was not registered using /regsvr32.

An out-of-process provider was not registered using /regserver.

$WbemErrorEnum.Add("wbemErrInitializationFailure", Dec(0x80041014))

$WbemErrorEnum.Add("wbemErrTransportFailure", Dec(0x80041015))

$WbemErrorEnum.Add("wbemErrInvalidOperation", Dec(0x80041016))

$WbemErrorEnum.Add("wbemErrInvalidQuery", Dec(0x80041017))

$WbemErrorEnum.Add("wbemErrInvalidQueryType", Dec(0x80041018))

$WbemErrorEnum.Add("wbemErrAlreadyExists", Dec(0x80041019))

$WbemErrorEnum.Add("wbemErrOverrideNotAllowed", Dec(0x8004101A))

$WbemErrorEnum.Add("wbemErrPropagatedQualifier", Dec(0x8004101B))

$WbemErrorEnum.Add("wbemErrPropagatedProperty", Dec(0x8004101C))

$WbemErrorEnum.Add("wbemErrUnexpected", Dec(0x8004101D))

$WbemErrorEnum.Add("wbemErrIllegalOperation", Dec(0x8004101E))

$WbemErrorEnum.Add("wbemErrCannotBeKey", Dec(0x8004101F))

$WbemErrorEnum.Add("wbemErrIncompleteClass", Dec(0x80041020))

$WbemErrorEnum.Add("wbemErrInvalidSyntax", Dec(0x80041021))

$WbemErrorEnum.Add("wbemErrNondecoratedObject", Dec(0x80041022))

$WbemErrorEnum.Add("wbemErrReadOnly", Dec(0x80041023))

$WbemErrorEnum.Add("wbemErrProviderNotCapable", Dec(0x80041024))

$WbemErrorEnum.Add("wbemErrClassHasChildren", Dec(0x80041025))

$WbemErrorEnum.Add("wbemErrClassHasInstances", Dec(0x80041026))

$WbemErrorEnum.Add("wbemErrQueryNotImplemented", Dec(0x80041027))

$WbemErrorEnum.Add("wbemErrIllegalNull", Dec(0x80041028))

$WbemErrorEnum.Add("wbemErrInvalidQualifierType", Dec(0x80041029))

$WbemErrorEnum.Add("wbemErrInvalidPropertyType", Dec(0x8004102A))

$WbemErrorEnum.Add("wbemErrValueOutOfRange", Dec(0x8004102B))

$WbemErrorEnum.Add("wbemErrCannotBeSingleton", Dec(0x8004102C))

$WbemErrorEnum.Add("wbemErrInvalidCimType", Dec(0x8004102D))

$WbemErrorEnum.Add("wbemErrInvalidMethod", Dec(0x8004102E))

$WbemErrorEnum.Add("wbemErrInvalidMethodParameters", Dec(0x8004102F))

$WbemErrorEnum.Add("wbemErrSystemProperty", Dec(0x80041030))

$WbemErrorEnum.Add("wbemErrInvalidProperty", Dec(0x80041031))

$WbemErrorEnum.Add("wbemErrCallCancelled", Dec(0x80041032))

$WbemErrorEnum.Add("wbemErrShuttingDown", Dec(0x80041033))

$WbemErrorEnum.Add("wbemErrPropagatedMethod", Dec(0x80041034))

$WbemErrorEnum.Add("wbemErrUnsupportedParameter", Dec(0x80041035))

$WbemErrorEnum.Add("wbemErrMissingParameter", Dec(0x80041036))

$WbemErrorEnum.Add("wbemErrInvalidParameterId", Dec(0x80041037))

$WbemErrorEnum.Add("wbemErrNonConsecutiveParameterIds", Dec(0x80041038))

$WbemErrorEnum.Add("wbemErrParameterIdOnRetval", Dec(0x80041039))

$WbemErrorEnum.Add("wbemErrInvalidObjectPath", Dec(0x8004103A))

wbemErrOutOfDiskSpace
2147749947



    Windows XP/2000/NT:  Disk is out of space.

    Windows Server 2003:  Disk is out of space or the 4 GB limit on WMI repository (CIM repository) size is reached.

$WbemErrorEnum.Add("wbemErrBufferTooSmall", Dec(0x8004103C))

$WbemErrorEnum.Add("wbemErrUnsupportedPutExtension", Dec(0x8004103D))

$WbemErrorEnum.Add("wbemErrUnknownObjectType", Dec(0x8004103E))

$WbemErrorEnum.Add("wbemErrUnknownPacketType", Dec(0x8004103F))

$WbemErrorEnum.Add("wbemErrMarshalVersionMismatch", Dec(0x80041040))

$WbemErrorEnum.Add("wbemErrMarshalInvalidSignature", Dec(0x80041041))

$WbemErrorEnum.Add("wbemErrInvalidQualifier", Dec(0x80041042))

$WbemErrorEnum.Add("wbemErrInvalidDuplicateParameter", Dec(0x80041043))

$WbemErrorEnum.Add("wbemErrTooMuchData", Dec(0x80041044))

$WbemErrorEnum.Add("wbemErrServerTooBusy", Dec(0x80041045))

$WbemErrorEnum.Add("wbemErrInvalidFlavor", Dec(0x80041046))

$WbemErrorEnum.Add("wbemErrCircularReference", Dec(0x80041047))

wbemErrUnsupportedClassUpdate
2147749960



The specified class is not supported.

$WbemErrorEnum.Add("wbemErrCannotChangeKeyInheritance", Dec(0x80041049))

$WbemErrorEnum.Add("wbemErrCannotChangeIndexInheritance", Dec(0x80041050))

$WbemErrorEnum.Add("wbemErrTooManyProperties", Dec(0x80041051))

$WbemErrorEnum.Add("wbemErrUpdateTypeMismatch", Dec(0x80041052))

$WbemErrorEnum.Add("wbemErrUpdateOverrideNotAllowed", Dec(0x80041053))

$WbemErrorEnum.Add("wbemErrUpdatePropagatedMethod", Dec(0x80041054))

$WbemErrorEnum.Add("wbemErrMethodNotImplemented", Dec(0x80041055))

$WbemErrorEnum.Add("wbemErrMethodDisabled", Dec(0x80041056))

$WbemErrorEnum.Add("wbemErrRefresherBusy", Dec(0x80041057))

$WbemErrorEnum.Add("wbemErrUnparsableQuery", Dec(0x80041058))

$WbemErrorEnum.Add("wbemErrNotEventClass", Dec(0x80041059))

$WbemErrorEnum.Add("wbemErrMissingGroupWithin", Dec(0x8004105A))

$WbemErrorEnum.Add("wbemErrMissingAggregationList", Dec(0x8004105B))

$WbemErrorEnum.Add("wbemErrPropertyNotAnObject", Dec(0x8004105C))

$WbemErrorEnum.Add("wbemErrAggregatingByObject", Dec(0x8004105D))

$WbemErrorEnum.Add("wbemErrUninterpretableProviderQuery", Dec(0x8004105F))

$WbemErrorEnum.Add("wbemErrBackupRestoreWinmgmtRunning", Dec(0x80041060))

$WbemErrorEnum.Add("wbemErrQueueOverflow", Dec(0x80041061))

$WbemErrorEnum.Add("wbemErrPrivilegeNotHeld", Dec(0x80041062))

$WbemErrorEnum.Add("wbemErrInvalidOperator", Dec(0x80041063))

$WbemErrorEnum.Add("wbemErrLocalCredentials", Dec(0x80041064))

$WbemErrorEnum.Add("wbemErrCannotBeAbstract", Dec(0x80041065))

$WbemErrorEnum.Add("wbemErrAmendedObject", Dec(0x80041066))

$WbemErrorEnum.Add("wbemErrClientTooSlow", Dec(0x80041067))

$WbemErrorEnum.Add("wbemErrNullSecurityDescriptor", Dec(0x80041068))

$WbemErrorEnum.Add("wbemErrTimeout", Dec(0x80041069))

$WbemErrorEnum.Add("wbemErrInvalidAssociation", Dec(0x8004106A))

$WbemErrorEnum.Add("wbemErrAmbiguousOperation", Dec(0x8004106B))

$WbemErrorEnum.Add("wbemErrQuotaViolation", Dec(0x8004106C))

$WbemErrorEnum.Add("wbemErrTransactionConflict", Dec(0x8004106D))

$WbemErrorEnum.Add("wbemErrForcedRollback", Dec(0x8004106E))

$WbemErrorEnum.Add("wbemErrUnsupportedLocale", Dec(0x8004106F))

$WbemErrorEnum.Add("wbemErrHandleOutOfDate", Dec(0x80041070))

$WbemErrorEnum.Add("wbemErrConnectionFailed", Dec(0x80041071))

$WbemErrorEnum.Add("wbemErrInvalidHandleRequest", Dec(0x80041072))

$WbemErrorEnum.Add("wbemErrPropertyNameTooWide", Dec(0x80041073))

$WbemErrorEnum.Add("wbemErrClassNameTooWide", Dec(0x80041074))

$WbemErrorEnum.Add("wbemErrMethodNameTooWide", Dec(0x80041075))

$WbemErrorEnum.Add("wbemErrQualifierNameTooWide", Dec(0x80041076))

$WbemErrorEnum.Add("wbemErrRerunCommand", Dec(0x80041077))

$WbemErrorEnum.Add("wbemErrDatabaseVerMismatch", Dec(0x80041078))

$WbemErrorEnum.Add("wbemErrVetoDelete", Dec(0x8004107A))

$WbemErrorEnum.Add("wbemErrVetoPut", Dec(0x8004107A))

$WbemErrorEnum.Add("wbemErrInvalidLocale", Dec(0x80041080))

$WbemErrorEnum.Add("wbemErrProviderSuspended", Dec(0x80041081))

$WbemErrorEnum.Add("wbemErrSynchronizationRequired", Dec(0x80041082))

$WbemErrorEnum.Add("wbemErrNoSchema", Dec(0x80041083))

$WbemErrorEnum.Add("wbemErrProviderAlreadyRegistered", Dec(0x80041084))

$WbemErrorEnum.Add("wbemErrProviderNotRegistered", Dec(0x80041085))

$WbemErrorEnum.Add("wbemErrFatalTransportError", Dec(0x80041086))

$WbemErrorEnum.Add("wbemErrEncryptedConnectionRequired", Dec(0x80041087))

$WbemErrorEnum.Add("WBEM_E_PROVIDER_TIMED_OUT", Dec(0x80041088))

$WbemErrorEnum.Add("WBEM_E_NO_KEY", Dec(0x80041089))

$WbemErrorEnum.Add("WBEM_E_PROVIDER_DISABLED", Dec(0x8004108A))

$WbemErrorEnum.Add("wbemErrRegistrationTooBroad", Dec(0x80042001))

$WbemErrorEnum.Add("wbemErrRegistrationTooPrecise", Dec(0x80042002))

$WbemErrorEnum.Add("wbemErrTimedout", Dec(0x80043001))

$WbemErrorEnum.Add("wbemErrResetToDefault", Dec(0x80043002))
#endregion

#region descriptions
$WbemErrorDescription.Add("wbemNoErr", "The call was successful.")

$WbemErrorDescription.Add("wbemErrFailed", "The call failed.")

$WbemErrorDescription.Add("wbemErrNotFound", "The object could not be found.")

$WbemErrorDescription.Add("wbemErrAccessDenied", "The current user does not have permission to perform the action.")

$WbemErrorDescription.Add("wbemErrProviderFailure", "The provider has failed at some time other than during initialization.")

$WbemErrorDescription.Add("wbemErrTypeMismatch", "A type mismatch occurred.")

$WbemErrorDescription.Add("wbemErrOutOfMemory", "There was not enough memory for the operation.")

$WbemErrorDescription.Add("wbemErrInvalidContext", "The SWbemNamedValue object is not valid.")

$WbemErrorDescription.Add("wbemErrInvalidParameter", "One of the parameters to the call is not correct.")

$WbemErrorDescription.Add("wbemErrNotAvailable", "The resource, typically a remote server, is not currently available.")

$WbemErrorDescription.Add("wbemErrCriticalError", "An internal, critical, and unexpected error occurred. Report this error to Microsoft Technical Support.")

$WbemErrorDescription.Add("wbemErrInvalidStream", "One or more network packets were corrupted during a remote session.")

$WbemErrorDescription.Add("wbemErrNotSupported", "The feature or operation is not supported.")

$WbemErrorDescription.Add("wbemErrInvalidSuperclass", "The parent class specified is not valid.")

$WbemErrorDescription.Add("wbemErrInvalidNamespace", "The namespace specified could not be found.")

$WbemErrorDescription.Add("wbemErrInvalidObject", "The specified instance is not valid.")

$WbemErrorDescription.Add("wbemErrInvalidClass", "The specified class is not valid.")

$WbemErrorDescription.Add("wbemErrProviderNotFound", "A provider referenced in the schema does not have a corresponding registration.")

$WbemErrorDescription.Add("wbemErrInvalidProviderRegistration", "A provider referenced in the schema has an incorrect or incomplete registration. This error may be caused by a missing pragma namespace command in the MOF file used to register the provider, resulting in the provider being registered in the wrong WMI namespace. This error may also be caused by a corrupt repository, which may be fixed by deleting it and recompiling the MOF files.")

$WbemErrorDescription.Add("wbemErrProviderLoadFailure", "COM cannot locate a provider referenced in the schema. This error may be caused by any of the following:")

The provider is using a WMI DLL that does not match the .lib fileused when the provider was built.

The provider's DLL or any of the DLLs on which it depends is corrupt.

The provider failed to export DllRegisterServer.

An in-process provider was not registered using /regsvr32.

An out-of-process provider was not registered using /regserver.

$WbemErrorDescription.Add("wbemErrInitializationFailure", "A component, such as a provider, failed to initialize for internal reasons.")

$WbemErrorDescription.Add("wbemErrTransportFailure", "A networking error occurred, preventing normal operation.")

$WbemErrorDescription.Add("wbemErrInvalidOperation", "The requested operation is not valid. This error usually applies to invalid attempts to delete classes or properties.")

$WbemErrorDescription.Add("wbemErrInvalidQuery", "The requested operation is not valid. This error usually applies to invalid attempts to delete classes or properties.")

$WbemErrorDescription.Add("wbemErrInvalidQueryType", "The requested query language is not supported.")

$WbemErrorDescription.Add("wbemErrAlreadyExists", "In a put operation, the wbemChangeFlagCreateOnly flag was specified, but the instance already exists.")

$WbemErrorDescription.Add("wbemErrOverrideNotAllowed", "It is not possible to perform the add operation on this qualifier because the owning object does not permit overrides.")

$WbemErrorDescription.Add("wbemErrPropagatedQualifier", "The user attempted to delete a qualifier that was not owned. The qualifier was inherited from a parent class.")

$WbemErrorDescription.Add("wbemErrPropagatedProperty", "The user attempted to delete a property that was not owned. The property was inherited from a parent class.")

$WbemErrorDescription.Add("wbemErrUnexpected", "The client made an unexpected and illegal sequence of calls, such as calling EndEnumeration before calling BeginEnumeration.")

$WbemErrorDescription.Add("wbemErrIllegalOperation", "The user requested an illegal operation, such as spawning a class from an instance.")

$WbemErrorDescription.Add("wbemErrCannotBeKey", "There was an illegal attempt to specify a key qualifier on a property that cannot be a key. The keys are specified in the class definition for an object, and cannot be altered on a per-instance basis.")

$WbemErrorDescription.Add("wbemErrIncompleteClass", "The current object is not a valid class definition. Either it is incomplete, or it has not been registered with WMI using SWbemObject.Put_.")

$WbemErrorDescription.Add("wbemErrInvalidSyntax", "The syntax of an input parameter is incorrect for the applicable data structure. For example, when a CIM datetime structure does not have the correct format when passed to SWbemDateTime.SetFileTime.")

$WbemErrorDescription.Add("wbemErrNondecoratedObject", "Reserved for future use.")

$WbemErrorDescription.Add("wbemErrReadOnly", "The property that you are attempting to modify is read-only.")

$WbemErrorDescription.Add("wbemErrProviderNotCapable", "The provider cannot perform the requested operation. This would include a query that is too complex, retrieving an instance, creating or updating a class, deleting a class, or enumerating a class.")

$WbemErrorDescription.Add("wbemErrClassHasChildren", "An attempt was made to make a change that would invalidate a subclass.")

$WbemErrorDescription.Add("wbemErrClassHasInstances", "An attempt has been made to delete or modify a class that has instances.")

$WbemErrorDescription.Add("wbemErrQueryNotImplemented", "Reserved for future use.")

$WbemErrorDescription.Add("wbemErrIllegalNull", "A value of Nothing was specified for a property that may not be Nothing, such as one that is marked by a Key, Indexed, or Not_Null qualifier.")

$WbemErrorDescription.Add("wbemErrInvalidQualifierType", "The CIM type specified for a property is not valid.")

$WbemErrorDescription.Add("wbemErrInvalidPropertyType", "The CIM type specified for a property is not valid.")

$WbemErrorDescription.Add("wbemErrValueOutOfRange", "The request was made with an out-of-range value, or is incompatible with the type.")

$WbemErrorDescription.Add("wbemErrCannotBeSingleton", "An illegal attempt was made to make a class singleton, such as when the class is derived from a non-singleton class.")

$WbemErrorDescription.Add("wbemErrInvalidCimType", "The CIM type specified is not valid.")

$WbemErrorDescription.Add("wbemErrInvalidMethod", "The requested method is not available.")

$WbemErrorDescription.Add("wbemErrInvalidMethodParameters", "The parameters provided for the method are not valid.")

$WbemErrorDescription.Add("wbemErrSystemProperty", "There was an attempt to get qualifiers on a system property.")

$WbemErrorDescription.Add("wbemErrInvalidProperty", "The property type is not recognized.")

$WbemErrorDescription.Add("wbemErrCallCancelled", "An asynchronous process has been canceled internally or by the user. Note that due to the timing and nature of the asynchronous operation the operation may not have been truly canceled.")

$WbemErrorDescription.Add("wbemErrShuttingDown", "The user has requested an operation while WMI is in the process of shutting down.")

$WbemErrorDescription.Add("wbemErrPropagatedMethod", "An attempt was made to reuse an existing method name from a parent class, and the signatures did not match.")

$WbemErrorDescription.Add("wbemErrUnsupportedParameter", "One or more parameter values, such as a query text, is too complex or unsupported. WMI is therefore requested to retry the operation with simpler parameters.")

$WbemErrorDescription.Add("wbemErrMissingParameter", "A parameter was missing from the method call.")

$WbemErrorDescription.Add("wbemErrInvalidParameterId", "A method parameter has an invalid ID qualifier.")

$WbemErrorDescription.Add("wbemErrNonConsecutiveParameterIds", "One or more of the method parameters have ID qualifiers that are out of sequence.")

$WbemErrorDescription.Add("wbemErrParameterIdOnRetval", "The return value for a method has an ID qualifier.")

$WbemErrorDescription.Add("wbemErrInvalidObjectPath", "The specified object path was invalid.")

wbemErrOutOfDiskSpace
2147749947



    Windows XP/2000/NT:  Disk is out of space.

    Windows Server 2003:  Disk is out of space or the 4 GB limit on WMI repository (CIM repository) size is reached.

$WbemErrorDescription.Add("wbemErrBufferTooSmall", "The supplied buffer was too small to hold all the objects in the enumerator or to read a string property.")

$WbemErrorDescription.Add("wbemErrUnsupportedPutExtension", "The provider does not support the requested put operation.")

$WbemErrorDescription.Add("wbemErrUnknownObjectType", "An object with an incorrect type or version was encountered during marshaling.")

$WbemErrorDescription.Add("wbemErrUnknownPacketType", "A packet with an incorrect type or version was encountered during marshaling.")

$WbemErrorDescription.Add("wbemErrMarshalVersionMismatch", "The packet has an unsupported version.")

$WbemErrorDescription.Add("wbemErrMarshalInvalidSignature", "The packet appears to be corrupted.")

$WbemErrorDescription.Add("wbemErrInvalidQualifier", "An attempt has been made to mismatch qualifiers, such as putting [key] on an object instead of a property.")

$WbemErrorDescription.Add("wbemErrInvalidDuplicateParameter", "A duplicate parameter has been declared in a CIM method.")

$WbemErrorDescription.Add("wbemErrTooMuchData", "Reserved for future use.")

$WbemErrorDescription.Add("wbemErrServerTooBusy", "A call to IWbemObjectSink::Indicate has failed. The provider may choose to refire the event.")

$WbemErrorDescription.Add("wbemErrInvalidFlavor", "The specified flavor was invalid.")

$WbemErrorDescription.Add("wbemErrCircularReference", "An attempt has been made to create a reference that is circular (for example, deriving a class from itself).")

wbemErrUnsupportedClassUpdate
2147749960



The specified class is not supported.

$WbemErrorDescription.Add("wbemErrCannotChangeKeyInheritance", "An attempt was made to change a key when instances or subclasses are already using the key.")

$WbemErrorDescription.Add("wbemErrCannotChangeIndexInheritance", "An attempt was made to change an index when instances or subclasses are already using the index.")

$WbemErrorDescription.Add("wbemErrTooManyProperties", "An attempt was made to create more properties than the current version of the class supports.")

$WbemErrorDescription.Add("wbemErrUpdateTypeMismatch", "A property was redefined with a conflicting type in a derived class.")

$WbemErrorDescription.Add("wbemErrUpdateOverrideNotAllowed", "An attempt was made in a derived class to override a non-overrideable qualifier.")

$WbemErrorDescription.Add("wbemErrUpdatePropagatedMethod", "A method was redeclared with a conflicting signature in a derived class.")

$WbemErrorDescription.Add("wbemErrMethodNotImplemented", "An attempt was made to execute a method not marked with [implemented] in any relevant class.")

$WbemErrorDescription.Add("wbemErrMethodDisabled", "An attempt was made to execute a method marked with [disabled].")

$WbemErrorDescription.Add("wbemErrRefresherBusy", "The refresher is busy with another operation.")

$WbemErrorDescription.Add("wbemErrUnparsableQuery", "The filtering query is syntactically invalid.")

$WbemErrorDescription.Add("wbemErrNotEventClass", "The FROM clause of a filtering query references a class that is not an event class (not derived from __Event).")

$WbemErrorDescription.Add("wbemErrMissingGroupWithin", "A GROUP BY clause was used without the corresponding GROUP WITHIN clause.")

$WbemErrorDescription.Add("wbemErrMissingAggregationList", "A GROUP BY clause was used. Aggregation on all properties is not supported.")

$WbemErrorDescription.Add("wbemErrPropertyNotAnObject", "Dot notation was used on a property that is not an embedded object.")

$WbemErrorDescription.Add("wbemErrAggregatingByObject", "A GROUP BY clause references a property that is an embedded object without using dot notation.")

$WbemErrorDescription.Add("wbemErrUninterpretableProviderQuery", "An event provider registration query ( __EventProviderRegistration) did not specify the classes for which events were provided.")

$WbemErrorDescription.Add("wbemErrBackupRestoreWinmgmtRunning", "An request was made to back up or restore the repository while WMI was using it.")

$WbemErrorDescription.Add("wbemErrQueueOverflow", "The asynchronous delivery queue overflowed due to the event consumer being too slow.")

$WbemErrorDescription.Add("wbemErrPrivilegeNotHeld", "The operation failed because the client did not have the necessary security privilege.")

$WbemErrorDescription.Add("wbemErrInvalidOperator", "The operator is not valid for this property type.")

$WbemErrorDescription.Add("wbemErrLocalCredentials", "The user specified a username, password or authority for a local connection. The user must use a blank username/password and rely on default security.")

$WbemErrorDescription.Add("wbemErrCannotBeAbstract", "The class was made abstract when its parent class is not abstract.")

$WbemErrorDescription.Add("wbemErrAmendedObject", "An amended object was put without the wbemFlagUseAmendedQualifiers flag being specified.")

$WbemErrorDescription.Add("wbemErrClientTooSlow", "Windows Server 2003 and Windows XP:  The client was not retrieving objects quickly enough from an enumeration. This constant is returned when a client creates an enumeration object but does not retrieve objects from the enumerator in a timely fashion, causing the enumerator's object caches to get backed up.")

$WbemErrorDescription.Add("wbemErrNullSecurityDescriptor", "Windows Server 2003 and Windows XP:  A null security descriptor was used.")

$WbemErrorDescription.Add("wbemErrTimeout", "Windows Server 2003 and Windows XP:  The operation timed out.")

$WbemErrorDescription.Add("wbemErrInvalidAssociation", "Windows Server 2003 and Windows XP:  The association being used is not valid.")

$WbemErrorDescription.Add("wbemErrAmbiguousOperation", "Windows Server 2003 and Windows XP:  The operation was ambiguous.")

$WbemErrorDescription.Add("wbemErrQuotaViolation", "Windows Server 2003 and Windows XP:  WMI is taking up too much memory. This could be caused either by low memory availability or excessive memory consumption by WMI.")

$WbemErrorDescription.Add("wbemErrTransactionConflict", "Windows Server 2003 and Windows XP:  The operation resulted in a transaction conflict.")

$WbemErrorDescription.Add("wbemErrForcedRollback", "Windows Server 2003 and Windows XP:  The transaction forced a rollback.")

$WbemErrorDescription.Add("wbemErrUnsupportedLocale", "Windows Server 2003 and Windows XP:  The locale used in the call is not supported.")

$WbemErrorDescription.Add("wbemErrHandleOutOfDate", "Windows Server 2003 and Windows XP:  The object handle is out of date.")

$WbemErrorDescription.Add("wbemErrConnectionFailed", "Windows Server 2003 and Windows XP:  Indicates that the connection to the SQL database failed.")

$WbemErrorDescription.Add("wbemErrInvalidHandleRequest", "Windows Server 2003 and Windows XP:  The handle request was invalid.")

$WbemErrorDescription.Add("wbemErrPropertyNameTooWide", "Windows Server 2003 and Windows XP:  The property name contains more than 255 characters.")

$WbemErrorDescription.Add("wbemErrClassNameTooWide", "Windows Server 2003 and Windows XP:  The class name contains more than 255 characters.")

$WbemErrorDescription.Add("wbemErrMethodNameTooWide", "Windows Server 2003 and Windows XP:  The method name contains more than 255 characters.")

$WbemErrorDescription.Add("wbemErrQualifierNameTooWide", "Windows Server 2003 and Windows XP:  The qualifier name contains more than 255 characters.")

$WbemErrorDescription.Add("wbemErrRerunCommand", "Windows Server 2003 and Windows XP:  Indicates that an SQL command should be rerun because there is a deadlock in SQL. This can be returned only when data is being stored in an SQL database.")

$WbemErrorDescription.Add("wbemErrDatabaseVerMismatch", "Windows Server 2003 and Windows XP:  The database version does not match the version that the repository driver understands.")

$WbemErrorDescription.Add("wbemErrVetoDelete", "Windows Server 2003 and Windows XP:  WMI cannot do the delete operation because the provider does not allow it.")

$WbemErrorDescription.Add("wbemErrVetoPut", "Windows Server 2003 and Windows XP:  WMI cannot do the put operation because the provider does not allow it.")

$WbemErrorDescription.Add("wbemErrInvalidLocale", "Windows Server 2003 and Windows XP:  The specified locale identifier was not valid for the operation.")

$WbemErrorDescription.Add("wbemErrProviderSuspended", "Windows Server 2003 and Windows XP:  The provider is suspended.")

$WbemErrorDescription.Add("wbemErrSynchronizationRequired", "Windows Server 2003 and Windows XP:  The object must be committed and retrieved again before the requested operation can succeed. This constant is returned when an object must be committed and re-retrieved to see the property value.")

$WbemErrorDescription.Add("wbemErrNoSchema", "Windows Server 2003 and Windows XP:  The operation cannot be completed because no schema is available.")

$WbemErrorDescription.Add("wbemErrProviderAlreadyRegistered", "Windows Server 2003 and Windows XP:  The provider registration cannot be done because the provider is already registered.")

$WbemErrorDescription.Add("wbemErrProviderNotRegistered", "Windows Server 2003 and Windows XP:  The provider for the requested data is not registered.")

$WbemErrorDescription.Add("wbemErrFatalTransportError", "Windows Server 2003 and Windows XP:  A fatal transport error occurred and other transport will not be attempted.")

$WbemErrorDescription.Add("wbemErrEncryptedConnectionRequired", "Windows Server 2003 and Windows XP:  The client connection to WINMGMT must be encrypted for this operation. The IWbemServices proxy security settings should be adjusted and the operation retried.")

$WbemErrorDescription.Add("WBEM_E_PROVIDER_TIMED_OUT", "Windows Server 2003 and Windows XP:  A provider failed to report results within the specified timeout.")

$WbemErrorDescription.Add("WBEM_E_NO_KEY", "Windows Server 2003 and Windows XP:  User attempted to put an instance with no defined key.")

$WbemErrorDescription.Add("WBEM_E_PROVIDER_DISABLED", "Windows Server 2003 and Windows XP:  User attempted to register a provider instance but the COM server for the provider instance was unloaded.")

$WbemErrorDescription.Add("wbemErrRegistrationTooBroad", "Windows Server 2003 and Windows XP:  The provider registration overlaps with the system event domain.")

$WbemErrorDescription.Add("wbemErrRegistrationTooPrecise", "Windows Server 2003 and Windows XP:  A WITHIN clause was not used in this query.")

$WbemErrorDescription.Add("wbemErrTimedout", "Windows Server 2003 and Windows XP:  Automation-specific error.")

$WbemErrorDescription.Add("wbemErrResetToDefault", "Windows Server 2003 and Windows XP:  The user deleted an override default value for the current class. The default value for this property in the parent class has been reactivated. An automation-specific error.")
#endregion