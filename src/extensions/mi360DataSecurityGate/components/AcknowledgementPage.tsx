import * as React from 'react';

export interface IAcknowledgementPageProps {
  returnUrl: string;
  policyUrl: string;
  supportUrl: string;
  onAcknowledge: (returnUrl: string) => Promise<void>;
}

export const AcknowledgementPage: React.FC<IAcknowledgementPageProps> = ({
  returnUrl,
  policyUrl,
  supportUrl,
  onAcknowledge
}) => {
  const [isChecked, setIsChecked] = React.useState(false);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const headingRef = React.useRef<HTMLHeadingElement>(null);
  const errorRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    headingRef.current?.focus();
  }, []);

  React.useEffect(() => {
    if (error) {
      errorRef.current?.focus();
    }
  }, [error]);

  const handleSubmit = async (): Promise<void> => {
    if (!isChecked || isSubmitting) {
      return;
    }

    setIsSubmitting(true);
    setError(null);

    try {
      await onAcknowledge(returnUrl);
    } catch (e) {
      setError('We could not record your acknowledgement. Please try again.');
      setIsSubmitting(false);
    }
  };

  const errorMessageId = 'mi360AcknowledgementError';

  return (
    <div
      role="main"
      aria-live="polite"
      style={{
        maxWidth: 760,
        margin: '32px auto',
        padding: 24,
        background: '#ffffff',
        borderRadius: 8,
        border: '1px solid #e1e1e1',
        boxShadow: '0 2px 8px rgba(0,0,0,0.08)'
      }}
    >
      <h1 ref={headingRef} tabIndex={-1} style={{ marginTop: 0, fontSize: 28, outline: 'none' }}>
        MI360 Data Security Acknowledgement
      </h1>

      <div style={{ marginTop: 16 }}>
        <p>
          To continue using MI360, you must acknowledge the MI360 Data Security requirements and agree to comply with the applicable ZF Data Security policy.
        </p>
      </div>

      <div style={{ marginTop: 16 }}>
        <p>
          Policy reference:{' '}
          <a href={policyUrl} target="_blank" rel="noopener noreferrer">
            ZF Data Security Policy
          </a>
        </p>
      </div>

      <div style={{ marginTop: 16 }}>
        <p>
          Selecting the checkbox confirms that you agree to comply with the policy. It does not prove that you read or understood every detail of the full policy.
        </p>
      </div>

      {error && (
        <div
          ref={errorRef}
          id={errorMessageId}
          role="alert"
          tabIndex={-1}
          style={{ marginTop: 16, padding: 12, color: '#a4262c', background: '#fde7e9', outline: 'none' }}
        >
          {error}
        </div>
      )}

      <label style={{ display: 'flex', gap: 8, marginTop: 20, alignItems: 'flex-start' }}>
        <input
          type="checkbox"
          checked={isChecked}
          disabled={isSubmitting}
          aria-describedby={error ? errorMessageId : undefined}
          onChange={(event) => setIsChecked(event.currentTarget.checked)}
        />
        <span>I acknowledge the MI360 Data Security requirements and agree to comply with the referenced ZF Data Security policy.</span>
      </label>

      <div style={{ marginTop: 24 }}>
        <button
          type="button"
          disabled={!isChecked || isSubmitting}
          aria-describedby={error ? errorMessageId : undefined}
          onClick={handleSubmit}
          style={{ padding: '10px 18px', cursor: isChecked && !isSubmitting ? 'pointer' : 'not-allowed' }}
        >
          {isSubmitting ? 'Submitting...' : error ? 'Try again' : 'I acknowledge'}
        </button>
      </div>

      <div style={{ marginTop: 24 }}>
        <p>
          If you are having trouble, contact the{' '}
          <a href={supportUrl}>MI360 support route</a>.
        </p>
      </div>
    </div>
  );
};

