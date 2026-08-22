/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { Card } from '../ui/BaseComponents';
import { Sparkles, AlertCircle, TrendingUp } from 'lucide-react';

export const AIInsights: React.FC = () => {
  const insights = [
    {
      id: '1',
      type: 'positive',
      text: 'Your net worth increased by 12% this month. Great job on saving!',
      icon: TrendingUp,
      color: 'text-emerald-400',
      bg: 'bg-emerald-400/10',
    },
    {
      id: '2',
      type: 'warning',
      text: 'Maize Production Q4 is nearing maturity. Plan your reinvestment.',
      icon: AlertCircle,
      color: 'text-emerald-400',
      bg: 'bg-emerald-400/10',
    },
    {
      id: '3',
      type: 'info',
      text: 'Consider diversifying into Sukuk to reduce portfolio volatility.',
      icon: Sparkles,
      color: 'text-teal-400',
      bg: 'bg-teal-400/10',
    }
  ];

  return (
    <Card className="bg-slate-900 border-slate-800">
      <div className="flex items-center gap-2 mb-6">
        <Sparkles className="text-teal-400 w-5 h-5" />
        <h3 className="text-subheading font-bold text-white uppercase">AI Insights</h3>
      </div>
      <div className="space-y-4">
        {insights.map((insight) => (
          <div key={insight.id} className={`p-4 rounded-xl ${insight.bg} flex gap-4 items-start border border-white/5 shadow-inner`}>
            <div className={`${insight.color} mt-0.5`}>
              <insight.icon size={18} />
            </div>
            <p className="text-body-sm text-slate-300 leading-relaxed">{insight.text}</p>
          </div>
        ))}
      </div>
    </Card>
  );
};
