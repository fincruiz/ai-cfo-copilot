"use client";
import { useState } from "react";
import { RotateCcw } from "lucide-react";
import { Button } from "@/components/ui/button";
import { ConfirmDialog } from "@/components/ui/confirm-dialog";
import { workspaceService, type ResetScope } from "@/services/workspace-service";
import { getApiErrorMessage } from "@/lib/api";

export function ModuleResetButton({ scope, label, description, onReset }: { scope: ResetScope; label: string; description: string; onReset?: () => void }) {
  const [open,setOpen]=useState(false); const [loading,setLoading]=useState(false); const [error,setError]=useState("");
  async function reset(){ setLoading(true); setError(""); try { await workspaceService.resetScope(scope); setOpen(false); onReset?.(); if(!onReset) window.location.reload(); } catch(e){ setError(getApiErrorMessage(e)); } finally { setLoading(false); } }
  return <><div className="flex flex-col items-end gap-1"><Button variant="outline" size="sm" onClick={()=>setOpen(true)}><RotateCcw className="size-4"/>{label}</Button>{error?<span className="max-w-64 text-right text-xs text-destructive">{error}</span>:null}</div><ConfirmDialog open={open} title={label+"?"} description={description} confirmLabel={label} destructive onCancel={()=>setOpen(false)} onConfirm={()=>void reset()} loading={loading}/></>;
}
