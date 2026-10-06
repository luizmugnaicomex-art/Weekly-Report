import React from 'react';

interface NumberInputProps {
  label: string;
  value: number;
  onChange: (value: number) => void;
  suffix?: string;
}

export const NumberInput: React.FC<NumberInputProps> = ({
  label,
  value,
  onChange,
  suffix,
}) => (
  <label className="flex flex-col gap-1 text-xs md:text-sm text-slate-600">
    <span className="font-medium">{label}</span>
    <div className="flex items-center gap-2">
      <input
        type="number"
        value={value}
        onChange={(e) => onChange(Number(e.target.value))}
        className="w-full rounded-lg border border-slate-200 px-2 py-1.5 text-sm text-slate-900 focus:outline-none focus:ring-2 focus:ring-red-500/70 focus:border-red-500/60 transition-all"
      />
      {suffix && (
        <span className="text-xs text-slate-500 whitespace-nowrap">
          {suffix}
        </span>
      )}
    </div>
  </label>
);