// src/components/TextAreaField.tsx
import React from "react";

interface TextAreaFieldProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  minRows?: number;
}

export const TextAreaField: React.FC<TextAreaFieldProps> = ({
  label,
  value,
  onChange,
  placeholder,
  minRows = 5,
}) => {
  return (
    <div className="flex flex-col gap-1.5 text-xs md:text-sm text-slate-700">
      <label className="font-medium text-slate-700">{label}</label>
      <textarea
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        rows={minRows}
        className="
          w-full rounded-lg border border-slate-200 bg-white
          px-3 py-2 text-sm text-slate-900 leading-relaxed
          shadow-sm focus:outline-none focus:ring-2 focus:ring-red-500/30
          focus:border-red-500/60 resize-none
          print:shadow-none print:bg-white print:text-[11px]
        "
      />
    </div>
  );
};