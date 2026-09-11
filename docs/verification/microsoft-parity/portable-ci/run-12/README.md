# Hosted regression at 433a199

Run [34492410734](https://github.com/NeoMei/WPSComposer/actions/runs/34492410734) succeeds at `433a199dd604ffbf42e827fb85e9a807a6626eed` on all four jobs. Ubuntu Python 3.9 and 3.12 each pass 5,196 tests with 44 skips. Windows Python 3.9 and 3.12 each pass 5,142 tests with 98 skips. Raw logs, JUnit, metadata and hashes are retained here.

This verifies the POSIX identity-path repair after run-11's 79 Windows failures, plus the bounded launcher, test isolation and list argument candidate. The additional Windows skips include POSIX subprocess-helper tests; pure protocol and path contracts remain covered. Opt-in native Office tests are not executed by these runners. Hosted Windows success does not replace Windows Office/WPS native, UI, installed-bundle or full parity certification.
