
import React from 'react';

interface CardProps {
  title: string;
  subtitle?: string;
  children: React.ReactNode;
  action?: React.ReactNode;
}

export const Card: React.FC<CardProps> = ({ title, subtitle, children, action }) => (
  <section className="bg-white/90 border border-slate-200 rounded-2xl shadow-sm p-4 md:p-5 flex flex-col gap-3 break-inside-avoid h-full">
    <header className="flex items-center justify-between gap-2">
      <div>
        <h2 className="text-sm md:text-base font-semibold text-slate-900">
          {title}
        </h2>
        {subtitle && (
          <p className="text-xs md:text-[13px] text-slate-500 mt-0.5">
            {subtitle}
          </p>
        )}
      </div>
      {action && (
        <div className="shrink-0">
          {action}
        </div>
      )}
    </header>
    <div className="flex-1">{children}</div>
  </section>
);
