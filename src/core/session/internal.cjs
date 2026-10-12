'use strict';

/**
 * Core-private binding between a SessionHandle and its transport driver.
 *
 * The driver carries open/close/forceSave/canSave. If the handle exposed it as a public field,
 * an adapter holding the handle could reach close() or the persistence barrier — exactly the two
 * things the M0 contract forbids it from touching (M0 3.1 forbidden edges, section 5). Keeping it
 * in a module-private WeakMap makes the prohibition structural: an adapter cannot reach the driver
 * without importing core/session, and the static gate proves no adapter does.
 */

const drivers = new WeakMap();

function bindDriver(handle, driver) {
  drivers.set(handle, driver);
  return handle;
}

function driverOf(handle) {
  const driver = drivers.get(handle);
  if (!driver) throw new Error('session handle has no bound driver');
  return driver;
}

module.exports = { bindDriver, driverOf };
