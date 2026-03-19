import { useState, useCallback } from 'react';

export type NotificationType = 'success' | 'error';

export interface Notification {
  type: NotificationType;
  message: string;
}

export function useNotifications() {
  const [notification, setNotification] = useState<Notification | null>(null);

  const showNotification = useCallback((type: NotificationType, message: string) => {
    setNotification({ type, message });
    // Auto-dismiss after 5 seconds
    setTimeout(() => {
      setNotification(prev => (prev?.message === message ? null : prev));
    }, 5000);
  }, []);

  const clearNotification = useCallback(() => {
    setNotification(null);
  }, []);

  return {
    notification,
    showNotification,
    clearNotification,
    setNotification
  };
}
