// =====================================================
// FILE: 02_Auth.gs
// =====================================================

const AUTH_SESSION_PREFIX = "SESSION_";
const AUTH_HASH_PREFIX = "sha256$";

function loginUser(email, password) {
  email = String(email || "").trim().toLowerCase();
  password = String(password || "").trim();

  if (!email || !password) {
    throw new Error("Debe escribir email y password.");
  }

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
    return user;
  }

  throw new Error("Email o password incorrecto.");
}

function loginUserForCompany(email, password, requestedCompanyId) {
  const user = loginUser(email, password);
  const requested = String(requestedCompanyId || "").trim().toUpperCase();

  if (
    requested &&
    String(user.role || "").trim().toUpperCase() !== "OWNER" &&
    String(user.companyId || "").trim().toUpperCase() !== requested
  ) {
    destroySession(user.sessionToken);
    throw new Error("Este login pertenece a otra compania.");
  }

  return user;
}

function ensureUserSecurityColumns_(sh) {
  let headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim();
    });

  ["PASSWORD_HASH", "PASSWORD_UPDATED_AT", "LAST_LOGIN_AT"].forEach(function(columnName) {
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

  if (requestedCompany && role !== "OWNER" && sessionCompany !== requestedCompany) {
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
