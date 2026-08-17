import React from 'react';

interface InputGroupProps {
  label: string;
  subLabel?: string;
  children: React.ReactNode;
  icon?: React.ReactNode;
}

export const InputGroup: React.FC<InputGroupProps> = ({ label, subLabel, children, icon }) => {
  return (
    <div className="mb-5">
      <div className="flex items-center gap-2 mb-2">
        {icon && <span className="text-blue-600">{icon}</span>}
        <label className="block text-sm font-semibold text-gray-700">
          {label}
        </label>
      </div>
      {children}
      {subLabel && <p className="mt-1 text-xs text-gray-500">{subLabel}</p>}
    </div>
  );
};