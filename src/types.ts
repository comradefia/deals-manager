/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface InboundRow {
  id: string;
  destination: string;
  isMasked: boolean;
  plannedMinutes: number;
  rpm: number; // Revenue Per Minute
}

export interface OutboundRow {
  id: string;
  destination: string;
  isMasked: boolean;
  plannedMinutes: number;
  cpm: number; // Cost Per Minute
}

export interface DealSummary {
  totalInboundMinutes: number;
  totalOutboundMinutes: number;
  totalPlannedRevenue: number;
  totalPlannedCost: number;
  netMargin: number;
  marginPercentage: number;
}
