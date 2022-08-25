[![CI](https://github.com/OuterlimitsTech/olt-dotnet-newtonsoft/actions/workflows/build.yml/badge.svg)](https://github.com/OuterlimitsTech/olt-dotnet-newtonsoft/actions/workflows/build.yml) [![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=OuterlimitsTech_olt-dotnet-newtonsoft&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=OuterlimitsTech_olt-dotnet-newtonsoft)

## [Newtonsoft.Json](https://www.nuget.org/packages/Newtonsoft.Json/) TypeCoverter, JsonConvert, & other classes targeting Netwonsoft.

### How to Use

| Utility/Item/Object                             | Description                               | Comments                                                                               |
| ----------------------------------------------- | ----------------------------------------- | -------------------------------------------------------------------------------------- |
| [OltEnumRouteConverter](#OltEnumRouteConverter) | Serialization of Enum in route parameters | [See Also](https://github.com/dotnet/aspnetcore/issues/4008)                           |
| [OltOnlyDateConverter](#OltOnlyDateConverter)   | Serialization of DateTime to yyyy-MM-dd   | Serializes the date without the time to allow frontend to parse into current timezones |

### OltEnumRouteConverter - TypeConverter

```csharp
[TypeConverter(typeof(OltEnumRouteConverter<DocumentType>))]
public enum DocumentType
{
    [EnumMember(Value = "passport")]
    Passport,

    [EnumMember(Value = "proof_of_address")]
    ProofOfAddress,
}
```

### OltOnlyDateConverter - JsonConverter that interits from IsoDateTimeConverter

```csharp
public class SampleModel
{
    [JsonConverter(typeof(OltOnlyDateConverter))]
    public DateTime BirthDate { get; set; }
}
```
