"use client";

import {
  createContext,
  useContext,
  useLayoutEffect,
  type ReactNode,
} from "react";
import type { ResolvedAccess } from "@/auth/access";
import type { PortalRole } from "@/lib/portal-users";
import { setPortalUserEmail } from "@/lib/portal-api";

export type AccessContextValue = {
  email: string;
  role: PortalRole;
  canEdit: boolean;
  canManageUsers: boolean;
};

const AccessContext = createContext<AccessContextValue>({
  email: "",
  role: "team",
  canEdit: false,
  canManageUsers: false,
});

export function AccessProvider({
  access,
  children,
}: {
  access: ResolvedAccess;
  children: ReactNode;
}) {
  useLayoutEffect(() => {
    setPortalUserEmail(access.email);
    return () => {
      setPortalUserEmail("");
    };
  }, [access.email]);

  return (
    <AccessContext.Provider
      value={{
        email: access.email,
        role: access.role,
        canEdit: access.canEdit,
        canManageUsers: access.canManageUsers,
      }}
    >
      {children}
    </AccessContext.Provider>
  );
}

export function useAccess(): AccessContextValue {
  return useContext(AccessContext);
}
