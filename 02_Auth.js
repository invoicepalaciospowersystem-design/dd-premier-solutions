// =====================================================
// FILE: 02_Auth.gs
// =====================================================

const AUTH_SESSION_PREFIX = "SESSION_";
const AUTH_HASH_PREFIX = "sha256$";
const AUTH_LOGIN_FAIL_PREFIX = "LOGIN_FAIL_";

function loginUser(email, password) {
  email = String(email || "").trim().toLowerCase();
  password = String(password || "").trim();

  if (!email || !password) {
    throw new Error("Debe escribir email y password.");
  }

  assertLoginNotLocked_(email);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_USERS);

  if (!sh) throw new Error("No existe la hoja USERS.");

  const headers = ensureUserSecurityColumns_(sh);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error("La hoja USERS no tiene usuarios.");

  const normalizedHeaders = headers.map(function(h) {
    return String(h).trim().toUpperCase();
  });

  const idxEmail = normalizedHeaders.indexOf("EMAIL");
  const idxName = normalizedHeaders.indexOf("NAME");
  const idxPassword = normalizedHeaders.indexOf("PASSWORD");
  const idxPasswordHash = normalizedHeaders.indexOf("PASSWORD_HASH");
  const idxRole = normalizedHeaders.indexOf("ROLE");
  const idxCompany = normalizedHeaders.indexOf("COMPANY_ID");
  const idxActive = normalizedHeaders.indexOf("ACTIVE");

  if ([idxEmail, idxName, idxRole, idxCompany, idxActive].indexOf(-1) !== -1) {
    throw new Error("USERS debe tener EMAIL, NAME, ROLE, COMPANY_ID y ACTIVE.");
  }

  if (idxPassword === -1 && idxPasswordHash === -1) {
    throw new Error("USERS debe tener PASSWORD o PASSWORD_HASH.");
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowEmail = String(row[idxEmail] || "").trim().toLowerCase();
    const active = String(row[idxActive] || "").trim().toUpperCase();

    if (rowEmail !== email) continue;
    if (active !== "YES") throw new Error("Usuario inactivo.");
    if (!verifyUserPassword_(sh, i + 1, headers, row, password)) break;

    const user = {
      success: true,
      email: rowEmail,
      name: String(row[idxName] || "").trim(),
      role: String(row[idxRole] || "").trim().toUpperCase(),
      companyId: String(row[idxCompany] || "").trim().toUpperCase()
    };

    const session = createSession_(user);
    user.sessionToken = session.token;
    user.sessionExpiresAt = session.expiresAt;
    clearLoginFailures_(email);
    return user;
  }

  recordLoginFailure_(email);
  throw new Error("Email o password incorrecto.");
}

function loginUserForCompany(email, password, requestedCompanyId, ownerOnly) {
  const user = loginUser(email, password);
  const requested = String(requestedCompanyId || "").trim().toUpperCase();
  const role = String(user.role || "").trim().toUpperCase();

  if (ownerOnly === true || String(ownerOnly || "").trim() === "true" || String(ownerOnly || "").trim() === "1") {
    if (!isGlobalOwnerRole_(role)) {
      destroySession(user.sessionToken);
      throw new Error("Este acceso es solo para OWNER.");
    }
  }

  if (
    requested &&
    !isGlobalOwnerRole_(role) &&
    String(user.companyId || "").trim().toUpperCase() !== requested
  ) {
    destroySession(user.sessionToken);
    throw new Error("Este login pertenece a otra compania.");
  }

  return user;
}

function isGlobalOwnerRole_(role) {
  role = String(role || "").trim().toUpperCase();
  return role === "OWNER" || role === "SYSTEM";
}

function ensureUserSecurityColumns_(sh) {
  let headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim();
    });

  ["PASSWORD_HASH", "PASSWORD_UPDATED_AT", "LAST_LOGIN_AT", "DELETED_AT", "DELETED_BY"].forEach(function(columnName) {
    const hasColumn = headers.some(function(h) {
      return String(h || "").trim().toUpperCase() === columnName;
    });

    if (!hasColumn) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(columnName);
      headers.push(columnName);
    }
  });

  return sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim();
    });
}

function assertLoginNotLocked_(email) {
  const record = getLoginFailureRecord_(email);
  const lockedUntil = Number(record.lockedUntil || 0);

  if (lockedUntil && lockedUntil > Date.now()) {
    const minutes = Math.ceil((lockedUntil - Date.now()) / 60000);
    throw new Error("Demasiados intentos fallidos. Intente nuevamente en " + minutes + " minutos.");
  }
}

function recordLoginFailure_(email) {
  const props = PropertiesService.getScriptProperties();
  const key = getLoginFailureKey_(email);
  const record = getLoginFailureRecord_(email);
  const now = Date.now();
  const maxAttempts = Number(CFG.LOGIN_LOCK_MAX_ATTEMPTS || 5);
  const lockMinutes = Number(CFG.LOGIN_LOCK_MINUTES || 15);

  const next = {
    count: Number(record.count || 0) + 1,
    firstAt: record.firstAt || now,
    lockedUntil: 0
  };

  if (next.count >= maxAttempts) {
    next.lockedUntil = now + lockMinutes * 60 * 1000;
  }

  props.setProperty(key, JSON.stringify(next));

  if (next.lockedUntil) {
    try {
      notifySystemError_("AUTH_LOGIN_LOCKED", new Error("Login bloqueado por demasiados intentos fallidos."), {
        module: "AUTH",
        email: maskEmailForSecurity_(email),
        attempts: next.count,
        lockMinutes: lockMinutes
      });
    } catch (err) {
      Logger.log("WARN AUTH_LOGIN_LOCKED alert: " + err);
    }
  }
}

function clearLoginFailures_(email) {
  PropertiesService.getScriptProperties().deleteProperty(getLoginFailureKey_(email));
}

function getLoginFailureRecord_(email) {
  const raw = PropertiesService.getScriptProperties().getProperty(getLoginFailureKey_(email));
  if (!raw) return {};

  try {
    const parsed = JSON.parse(raw);
    const lockedUntil = Number(parsed.lockedUntil || 0);

    if (lockedUntil && lockedUntil < Date.now()) {
      clearLoginFailures_(email);
      return {};
    }

    return parsed || {};
  } catch (err) {
    clearLoginFailures_(email);
    return {};
  }
}

function getLoginFailureKey_(email) {
  return AUTH_LOGIN_FAIL_PREFIX + hashSessionToken_(String(email || "").trim().toLowerCase());
}

function maskEmailForSecurity_(email) {
  email = String(email || "").trim().toLowerCase();
  const parts = email.split("@");
  if (parts.length !== 2) return email ? "***" : "";

  const local = parts[0];
  const domainParts = parts[1].split(".");
  const visibleLocal = local.length <= 2 ? local.charAt(0) : local.slice(0, 2);
  const visibleDomain = domainParts[0] ? domainParts[0].charAt(0) + "***" : "***";
  const suffix = domainParts.length > 1 ? "." + domainParts.slice(1).join(".") : "";

  return visibleLocal + "***@" + visibleDomain + suffix;
}

function verifyUserPassword_(sh, rowNumber, headers, row, password) {
  const normalizedHeaders = headers.map(function(h) {
    return String(h).trim().toUpperCase();
  });

  const idxPassword = normalizedHeaders.indexOf("PASSWORD");
  const idxPasswordHash = normalizedHeaders.indexOf("PASSWORD_HASH");
  const idxUpdatedAt = normalizedHeaders.indexOf("PASSWORD_UPDATED_AT");
  const idxLastLogin = normalizedHeaders.indexOf("LAST_LOGIN_AT");

  const storedHash = idxPasswordHash >= 0 ? String(row[idxPasswordHash] || "").trim() : "";
  const legacyPassword = idxPassword >= 0 ? String(row[idxPassword] || "").trim() : "";

  if (storedHash && verifyPasswordHash_(password, storedHash)) {
    if (idxLastLogin >= 0) sh.getRange(rowNumber, idxLastLogin + 1).setValue(new Date());
    return true;
  }

  if (legacyPassword && legacyPassword === password) {
    if (idxPasswordHash >= 0) {
      sh.getRange(rowNumber, idxPasswordHash + 1).setValue(hashPassword_(password));
    }
    if (idxPassword >= 0) sh.getRange(rowNumber, idxPassword + 1).setValue("");
    if (idxUpdatedAt >= 0) sh.getRange(rowNumber, idxUpdatedAt + 1).setValue(new Date());
    if (idxLastLogin >= 0) sh.getRange(rowNumber, idxLastLogin + 1).setValue(new Date());
    return true;
  }

  return false;
}

function hashPassword_(password) {
  const salt = Utilities.getUuid().replace(/-/g, "");
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + ":" + String(password || ""),
    Utilities.Charset.UTF_8
  );

  return AUTH_HASH_PREFIX + salt + "$" + Utilities.base64EncodeWebSafe(digest);
}

function verifyPasswordHash_(password, storedHash) {
  storedHash = String(storedHash || "").trim();
  if (storedHash.indexOf(AUTH_HASH_PREFIX) !== 0) return false;

  const parts = storedHash.split("$");
  if (parts.length !== 3 || !parts[1] || !parts[2]) return false;

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    parts[1] + ":" + String(password || ""),
    Utilities.Charset.UTF_8
  );

  return constantTimeEquals_(parts[2], Utilities.base64EncodeWebSafe(digest));
}

function constantTimeEquals_(a, b) {
  a = String(a || "");
  b = String(b || "");

  let diff = a.length ^ b.length;
  const max = Math.max(a.length, b.length);

  for (let i = 0; i < max; i++) {
    diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }

  return diff === 0;
}

function createSession_(user) {
  const token = Utilities.getUuid() + "-" + Utilities.getUuid();
  const now = Date.now();
  const ttlSeconds = Math.min(Number(CFG.SESSION_TTL_SECONDS || 21600), 21600);
  const expiresAt = now + (ttlSeconds * 1000);

  const payload = JSON.stringify({
    email: String(user.email || "").trim().toLowerCase(),
    name: String(user.name || "").trim(),
    role: String(user.role || "").trim().toUpperCase(),
    companyId: String(user.companyId || "").trim().toUpperCase(),
    createdAt: now,
    expiresAt: expiresAt
  });

  const key = getSessionStorageKey_(token);
  CacheService.getScriptCache().put(key, payload, ttlSeconds);
  PropertiesService.getScriptProperties().setProperty(key, payload);

  return {
    token: token,
    expiresAt: expiresAt
  };
}

function getSession_(sessionToken) {
  sessionToken = String(sessionToken || "").trim();
  if (!sessionToken) return null;

  const key = getSessionStorageKey_(sessionToken);
  let payload = CacheService.getScriptCache().get(key);

  if (!payload) {
    payload = PropertiesService.getScriptProperties().getProperty(key);
  }

  if (!payload) return null;

  try {
    const session = JSON.parse(payload);

    if (!session.expiresAt || Number(session.expiresAt) < Date.now()) {
      PropertiesService.getScriptProperties().deleteProperty(key);
      CacheService.getScriptCache().remove(key);
      return null;
    }

    return session;
  } catch (err) {
    PropertiesService.getScriptProperties().deleteProperty(key);
    CacheService.getScriptCache().remove(key);
    return null;
  }
}

function requireSession_(sessionToken, allowedRoles, companyId) {
  const session = getSession_(sessionToken);
  if (!session) throw new Error("Sesion expirada. Inicie sesion nuevamente.");

  const role = String(session.role || "").trim().toUpperCase();
  const roles = (allowedRoles || []).map(function(r) {
    return String(r || "").trim().toUpperCase();
  }).filter(Boolean);

  if (roles.length && roles.indexOf(role) === -1) {
    throw new Error("No autorizado para esta accion.");
  }

  const requestedCompany = String(companyId || "").trim().toUpperCase();
  const sessionCompany = String(session.companyId || "").trim().toUpperCase();

  if (requestedCompany && !isGlobalOwnerRole_(role) && sessionCompany !== requestedCompany) {
    throw new Error("No autorizado para esta compania.");
  }

  return session;
}

function destroySession(sessionToken) {
  sessionToken = String(sessionToken || "").trim();
  if (!sessionToken) return true;

  const key = getSessionStorageKey_(sessionToken);
  CacheService.getScriptCache().remove(key);
  PropertiesService.getScriptProperties().deleteProperty(key);
  return true;
}

function cleanupExpiredSessions_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const now = Date.now();
  let removed = 0;

  Object.keys(all).forEach(function(key) {
    if (key.indexOf(AUTH_SESSION_PREFIX) !== 0) return;

    try {
      const session = JSON.parse(all[key] || "{}");
      if (!session.expiresAt || Number(session.expiresAt) < now) {
        props.deleteProperty(key);
        CacheService.getScriptCache().remove(key);
        removed++;
      }
    } catch (err) {
      props.deleteProperty(key);
      CacheService.getScriptCache().remove(key);
      removed++;
    }
  });

  return removed;
}

function getSessionStorageKey_(sessionToken) {
  return AUTH_SESSION_PREFIX + hashSessionToken_(sessionToken);
}

function hashSessionToken_(sessionToken) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(sessionToken || ""),
    Utilities.Charset.UTF_8
  );

  return Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, "");
}
