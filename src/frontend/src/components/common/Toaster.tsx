"use client";

import { createContext, useContext, type ReactNode } from "react";
import { Toast, ToastBody, Toaster as FluentToaster, ToastTitle, useToastController } from "@fluentui/react-components";

type ToastIntent = "success" | "error" | "warning" | "info";
interface ToastOptions {
  title: string;
  body?: string;
  intent?: ToastIntent;
}

const toasterId = "wordembedded-toaster";
const ToastContext = createContext<((options: ToastOptions) => void) | undefined>(undefined);

export function AppToasterProvider({ children }: { children: ReactNode }) {
  const { dispatchToast } = useToastController(toasterId);

  const notify = ({ title, body, intent = "info" }: ToastOptions) => {
    dispatchToast(
      <Toast>
        <ToastTitle>{title}</ToastTitle>
        {body ? <ToastBody>{body}</ToastBody> : null}
      </Toast>,
      { intent, timeout: 5000 },
    );
  };

  return (
    <ToastContext.Provider value={notify}>
      {children}
      <FluentToaster toasterId={toasterId} position="bottom-end" />
    </ToastContext.Provider>
  );
}

export function useAppToast() {
  const context = useContext(ToastContext);
  if (!context) throw new Error("useAppToast must be used within AppToasterProvider");
  return context;
}
