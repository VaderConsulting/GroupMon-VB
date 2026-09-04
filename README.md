# GroupMon

VB.NET Windows service that watches Active Directory group membership and emails alerts when members are added or removed. Versioned trees from 0.9 through 2.0 include the GroupMonSvc worker, MSI setup projects, and companion Test / GroupStatus apps. Operators configure monitored group LDAP paths, SMTP settings, and exclusion lists so only meaningful membership changes raise notifications.

**Source last updated:** 2012-07-07

---

## Contents

- Multiple versioned trees (`0.9`, `1.0`, `1.1`, `1.2`, `2.0`)
- Service projects (`GroupMonSvc`)
- Test/status companion apps (`Test`, `GroupStatus`)

## Attribution and provenance

No third-party source attribution markers were identified during this review.

## Requirements

- Visual Studio 2005

## License

MIT. See `LICENSE`.
