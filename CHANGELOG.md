# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.0.9] - 2026-03-13

### Added
- Added package.json metadata fields (repository, author, bugs, homepage)
- Added security documentation warning

### Fixed
- Fixed quick-start.md documentation - replaced VBSInterpreter with VbsEngine

### Security
- Added security warnings to documentation regarding execution of untrusted scripts

## [0.0.8] - 2026-??-??

### Added
- Initial npm release
- VBScript interpreter engine
- Browser runtime with DOM integration
- Node.js runtime support
- MSScriptControl API compatibility

### Known Issues
- Performance tests may be skipped in certain environments
- Timeout detection can be bypassed with tight synchronous loops
